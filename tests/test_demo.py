"""Demo signup must keep separate identities, copy permissions, and CSRF checks."""
import io
import json
import os
import secrets
import time
import unittest
import uuid
from unittest.mock import patch
from app.db import DB, init_schema
from app.server import application


class Demo(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with DB() as db:
            init_schema(db)

    def call(self, path, method='GET', body=None, cookie='', csrf='', origin='http://10.0.0.25:8000'):
        response = {}
        data = json.dumps(body).encode() if body is not None else b''
        env = dict(PATH_INFO=path, REQUEST_METHOD=method, CONTENT_LENGTH=str(len(data)),
                   HTTP_COOKIE=cookie, HTTP_X_CSRF_TOKEN=csrf, HTTP_ORIGIN=origin, REMOTE_ADDR='127.0.0.1')
        env['wsgi.input'] = io.BytesIO(data)
        def start(status, headers):
            response['status'] = int(status.split()[0])
            response['headers'] = dict(headers)
        response['body'] = json.loads(b''.join(application(env, start)))
        return response

    def signup(self):
        email = uuid.uuid4().hex + '@demo.test'
        password = secrets.token_urlsafe(24)
        response = self.call('/api/demo/register', 'POST', dict(name='Demo planner', email=email,
                             password=password, admin=1))
        self.assertEqual(response['status'], 200, response)
        self.assertEqual(response['body']['admin'], 0)
        cookie = response['headers']['Set-Cookie'].split(';')[0]
        self.assertIn('HttpOnly', response['headers']['Set-Cookie'])
        return response['body'], cookie, password

    def test_demo_users_are_private_and_sharing_needs_no_admin(self):
        with patch.dict(os.environ, PLANNING_DEMO_MODE='1', PUBLIC_ORIGINS='http://10.0.0.25:8000'):
            a, ca, pw = self.signup()
            b, cb, _ = self.signup()
            copy = self.call('/api/copies', 'POST', {'name': 'Demo copy'}, ca, a['csrf'])
            self.assertEqual(copy['status'], 200)
            sid = copy['body']['id']
            self.assertEqual(self.call('/api/copies/'+sid+'/members', cookie=cb)['status'], 403)
            share = self.call('/api/copies/'+sid+'/members', 'POST', {'email':b['email'],'role':'editor'}, ca, a['csrf'])
            self.assertEqual(share['status'], 200)
            self.assertEqual(self.call('/api/copies/'+sid+'/members', cookie=cb)['body']['role'], 'editor')
            self.assertEqual(self.call('/api/copies', 'POST', {'name':'No token'}, ca)['status'], 403)
            self.assertEqual(self.call('/api/users', 'POST', {'name':'No admin'}, ca, a['csrf'])['status'], 403)
            duplicate = self.call('/api/demo/register', 'POST', {'name':'Impersonator','email':a['email'],'password':pw})
            self.assertEqual(duplicate['status'], 409)
            again = self.call('/api/login', 'POST', {'email':a['email'],'password':pw})
            self.assertEqual(again['body']['id'], a['id'])

    def test_signup_is_disabled_outside_demo_and_checks_origin(self):
        with patch.dict(os.environ, PLANNING_DEMO_MODE='0', PUBLIC_ORIGINS='http://10.0.0.25:8000'):
            self.assertFalse(self.call('/api/demo')['body']['enabled'])
            self.assertEqual(self.call('/api/demo/register', 'POST', {})['status'], 403)
        with patch.dict(os.environ, PLANNING_DEMO_MODE='1', PUBLIC_ORIGINS='http://10.0.0.25:8000'):
            self.assertEqual(self.call('/api/demo/register', 'POST', {}, origin='https://other.example')['status'], 403)

    def test_invalid_signup_does_not_create_account(self):
        with patch.dict(os.environ, PLANNING_DEMO_MODE='1', PUBLIC_ORIGINS='http://10.0.0.25:8000'):
            with DB() as db:
                before = db.one('SELECT COUNT(*) AS n FROM users')['n']
            for body in ({'name':'','email':'a@demo.test','password':secrets.token_urlsafe(24)},
                         {'name':'A','email':'invalid','password':secrets.token_urlsafe(24)},
                         {'name':'A','email':'a@demo.test','password':'short'}, []):
                self.assertEqual(self.call('/api/demo/register', 'POST', body)['status'], 400)
            with DB() as db:
                self.assertEqual(db.one('SELECT COUNT(*) AS n FROM users')['n'], before)
