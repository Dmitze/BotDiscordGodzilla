PRAGMA journal_mode=WAL;
PRAGMA synchronous=NORMAL;

-- User stars (favorites)
CREATE TABLE IF NOT EXISTS user_stars (
  user_id TEXT NOT NULL,
  file_id TEXT NOT NULL,
  name TEXT,
  tags TEXT,
  added_at INTEGER NOT NULL,
  PRIMARY KEY(user_id, file_id)
);
CREATE INDEX IF NOT EXISTS idx_user_stars_user ON user_stars(user_id);

-- Saved searches
CREATE TABLE IF NOT EXISTS saved_searches (
  user_id TEXT NOT NULL,
  name TEXT NOT NULL,
  query_json TEXT NOT NULL,
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL,
  PRIMARY KEY(user_id, name)
);
CREATE INDEX IF NOT EXISTS idx_saved_searches_user ON saved_searches(user_id);

-- Recent items
CREATE TABLE IF NOT EXISTS recent_items (
  user_id TEXT NOT NULL,
  file_id TEXT NOT NULL,
  name TEXT,
  opened_at INTEGER NOT NULL,
  snippet TEXT,
  PRIMARY KEY(user_id, file_id)
);
CREATE INDEX IF NOT EXISTS idx_recent_items_user ON recent_items(user_id);

-- Subscriptions
CREATE TABLE IF NOT EXISTS subscriptions (
  user_id TEXT NOT NULL,
  topic TEXT NOT NULL,
  criteria_json TEXT,
  created_at INTEGER NOT NULL,
  PRIMARY KEY(user_id, topic)
);
CREATE INDEX IF NOT EXISTS idx_subscriptions_user ON subscriptions(user_id);

-- Change events (дедуп по file_id+change_id+hash)
CREATE TABLE IF NOT EXISTS change_events (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  file_id TEXT NOT NULL,
  change_id TEXT NOT NULL,
  hash TEXT NOT NULL,
  occurred_at INTEGER NOT NULL,
  meta_json TEXT
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_change_events ON change_events(file_id, change_id, hash);
CREATE INDEX IF NOT EXISTS idx_change_events_time ON change_events(occurred_at);

-- Очередь уведомлений (коалесинг по окну времени)
CREATE TABLE IF NOT EXISTS notifications_queue (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id TEXT NOT NULL,
  topic TEXT NOT NULL,
  file_id TEXT NOT NULL,
  change_id TEXT NOT NULL,
  hash TEXT NOT NULL,
  window_start INTEGER NOT NULL,
  window_end INTEGER NOT NULL,
  status TEXT NOT NULL DEFAULT 'pending', -- pending|delivered|failed
  payload_json TEXT,
  created_at INTEGER NOT NULL,
  updated_at INTEGER NOT NULL,
  delivered_at INTEGER
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_notifications_window ON notifications_queue(user_id, topic, file_id, change_id, hash, window_start);
CREATE INDEX IF NOT EXISTS idx_notifications_status ON notifications_queue(status, window_end);

-- Дайджесты (daily/weekly)
CREATE TABLE IF NOT EXISTS digests (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id TEXT NOT NULL,
  period TEXT NOT NULL, -- daily|weekly
  window_start INTEGER NOT NULL,
  window_end INTEGER NOT NULL,
  payload_json TEXT NOT NULL,
  created_at INTEGER NOT NULL,
  delivered_at INTEGER
);
CREATE INDEX IF NOT EXISTS idx_digests_user_period ON digests(user_id, period, window_start);
