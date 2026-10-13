"""Additive SQLite schema for the v2 backend."""


SCHEMA = """
CREATE TABLE IF NOT EXISTS pp_meta (
    key TEXT PRIMARY KEY, value TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_users (
    id TEXT PRIMARY KEY, email TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
    password_hash TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_sessions (
    token_hash TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES pp_users(id) ON DELETE CASCADE,
    csrf TEXT NOT NULL, expires_at TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_login_attempts (
    key TEXT NOT NULL, attempted_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS pp_login_attempts_key_time ON pp_login_attempts(key, attempted_at);
CREATE TABLE IF NOT EXISTS pp_categories (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, created_by TEXT NOT NULL REFERENCES pp_users(id),
    created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_category_members (
    category_id TEXT NOT NULL REFERENCES pp_categories(id) ON DELETE CASCADE,
    user_id TEXT NOT NULL REFERENCES pp_users(id) ON DELETE CASCADE,
    role TEXT NOT NULL CHECK(role IN ('manager','member')),
    PRIMARY KEY(category_id, user_id)
);
CREATE TABLE IF NOT EXISTS pp_templates (
    id TEXT PRIMARY KEY, name TEXT NOT NULL, source_path TEXT NOT NULL,
    snapshot_path TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_copies (
    id TEXT PRIMARY KEY, name TEXT NOT NULL,
    category_id TEXT REFERENCES pp_categories(id) ON DELETE SET NULL,
    owner_id TEXT NOT NULL REFERENCES pp_users(id),
    template_id TEXT NOT NULL REFERENCES pp_templates(id),
    visibility TEXT NOT NULL CHECK(visibility IN ('private','category')),
    default_role TEXT NOT NULL CHECK(default_role IN ('editor','viewer')),
    parent_copy_id TEXT REFERENCES pp_copies(id), fork_revision INTEGER,
    input_revision INTEGER NOT NULL DEFAULT 0,
    calculated_revision INTEGER NOT NULL DEFAULT -1,
    created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_copy_shares (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    user_id TEXT NOT NULL REFERENCES pp_users(id) ON DELETE CASCADE,
    role TEXT NOT NULL CHECK(role IN ('editor','viewer')),
    PRIMARY KEY(copy_id, user_id)
);
CREATE TABLE IF NOT EXISTS pp_overrides (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    sheet INTEGER NOT NULL, row INTEGER NOT NULL, col INTEGER NOT NULL,
    value_json TEXT NOT NULL, version INTEGER NOT NULL,
    updated_by TEXT NOT NULL REFERENCES pp_users(id), updated_at TEXT NOT NULL,
    origin_copy_id TEXT, origin_revision INTEGER,
    PRIMARY KEY(copy_id, sheet, row, col)
);
CREATE TABLE IF NOT EXISTS pp_copy_baselines (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    sheet INTEGER NOT NULL, row INTEGER NOT NULL, col INTEGER NOT NULL,
    value_json TEXT NOT NULL,
    PRIMARY KEY(copy_id, sheet, row, col)
);
CREATE TABLE IF NOT EXISTS pp_ops (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    op_id TEXT NOT NULL, actor_id TEXT NOT NULL REFERENCES pp_users(id),
    payload_hash TEXT NOT NULL, ack_json TEXT NOT NULL, action TEXT NOT NULL,
    revision INTEGER NOT NULL, created_at TEXT NOT NULL,
    PRIMARY KEY(copy_id, op_id)
);
CREATE TABLE IF NOT EXISTS pp_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    op_id TEXT NOT NULL, revision INTEGER NOT NULL, at TEXT NOT NULL,
    actor_id TEXT NOT NULL REFERENCES pp_users(id), action TEXT NOT NULL,
    sheet INTEGER NOT NULL, row INTEGER NOT NULL, col INTEGER NOT NULL,
    before_json TEXT NOT NULL, after_json TEXT NOT NULL, version INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS pp_history_copy_revision ON pp_history(copy_id, revision, id);
CREATE TABLE IF NOT EXISTS pp_calc_jobs (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    revision INTEGER NOT NULL, status TEXT NOT NULL,
    attempts INTEGER NOT NULL DEFAULT 0, last_error TEXT,
    created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    PRIMARY KEY(copy_id, revision)
);
CREATE TABLE IF NOT EXISTS pp_results (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    sheet INTEGER NOT NULL, row INTEGER NOT NULL, col INTEGER NOT NULL,
    value_json TEXT NOT NULL, calculated_revision INTEGER NOT NULL,
    PRIMARY KEY(copy_id, sheet, row, col)
);
CREATE TABLE IF NOT EXISTS pp_events (
    sequence INTEGER PRIMARY KEY AUTOINCREMENT,
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    event_type TEXT NOT NULL, input_revision INTEGER NOT NULL,
    payload_json TEXT NOT NULL, created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS pp_events_copy_sequence ON pp_events(copy_id, sequence);
CREATE TABLE IF NOT EXISTS pp_source_overlays (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    source_id TEXT NOT NULL, revision INTEGER NOT NULL, payload_json TEXT NOT NULL,
    PRIMARY KEY(copy_id, source_id, revision)
);
CREATE TABLE IF NOT EXISTS pp_dashboard_views (
    copy_id TEXT NOT NULL REFERENCES pp_copies(id) ON DELETE CASCADE,
    user_id TEXT NOT NULL REFERENCES pp_users(id) ON DELETE CASCADE,
    sheet INTEGER NOT NULL, pivot_name TEXT NOT NULL,
    view_json TEXT NOT NULL, updated_at TEXT NOT NULL,
    PRIMARY KEY(copy_id, user_id, sheet, pivot_name)
);
CREATE TABLE IF NOT EXISTS pp_category_dashboard_views (
    category_id TEXT NOT NULL REFERENCES pp_categories(id) ON DELETE CASCADE,
    user_id TEXT NOT NULL REFERENCES pp_users(id) ON DELETE CASCADE,
    sheet INTEGER NOT NULL, pivot_name TEXT NOT NULL,
    view_json TEXT NOT NULL, updated_at TEXT NOT NULL,
    PRIMARY KEY(category_id, user_id, sheet, pivot_name)
);
CREATE TABLE IF NOT EXISTS pp_pivot_state (
    copy_id TEXT PRIMARY KEY REFERENCES pp_copies(id) ON DELETE CASCADE,
    state_json TEXT NOT NULL, revision INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS pp_legacy_events (
    event_id TEXT PRIMARY KEY
);
"""
