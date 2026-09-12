CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS import_batch (
    id uuid PRIMARY KEY,
    filename text NOT NULL,
    file_sha256 char(64) NOT NULL,
    uploaded_at timestamptz NOT NULL DEFAULT now(),
    inferred_start_date date,
    inferred_end_date date,
    confirmed_start_date date,
    confirmed_end_date date,
    row_count integer NOT NULL DEFAULT 0,
    status text NOT NULL,
    validation_result jsonb NOT NULL DEFAULT '{}'::jsonb,
    original_file_path text NOT NULL,
    archived_file_path text,
    synthetic_reconciliation boolean NOT NULL DEFAULT false,
    UNIQUE (file_sha256, original_file_path)
);

CREATE TABLE IF NOT EXISTS outage_record_version (
    id bigserial PRIMARY KEY,
    batch_id uuid NOT NULL REFERENCES import_batch(id) ON DELETE RESTRICT,
    source_sheet text NOT NULL,
    source_row integer NOT NULL,
    record_key text NOT NULL,
    content_hash char(64) NOT NULL,
    source_record_id text,
    event_id text,
    user_id text,
    work_order text,
    customer_code text,
    feeder_code text,
    outage_start timestamp NOT NULL,
    outage_end timestamp,
    city text,
    district text,
    station text,
    raw_payload jsonb NOT NULL,
    activated_at timestamptz,
    superseded_at timestamptz,
    superseded_by_batch_id uuid REFERENCES import_batch(id),
    rolled_back_at timestamptz,
    UNIQUE (batch_id, record_key)
);

CREATE INDEX IF NOT EXISTS outage_record_current_range_idx
    ON outage_record_version (outage_start, record_key)
    WHERE activated_at IS NOT NULL AND superseded_at IS NULL AND rolled_back_at IS NULL;
CREATE INDEX IF NOT EXISTS outage_record_batch_idx ON outage_record_version (batch_id);

CREATE TABLE IF NOT EXISTS batch_record_audit (
    batch_id uuid NOT NULL REFERENCES import_batch(id) ON DELETE RESTRICT,
    record_key text NOT NULL,
    content_hash char(64) NOT NULL,
    source_sheet text NOT NULL,
    source_row integer NOT NULL,
    outage_start timestamp NOT NULL,
    city text,
    version_id bigint REFERENCES outage_record_version(id) ON DELETE SET NULL,
    PRIMARY KEY (batch_id, record_key)
);

CREATE INDEX IF NOT EXISTS batch_record_audit_version_idx
    ON batch_record_audit (version_id);

CREATE TABLE IF NOT EXISTS mask_batch (
    id uuid PRIMARY KEY,
    filename text NOT NULL,
    file_sha256 char(64) NOT NULL,
    uploaded_at timestamptz NOT NULL DEFAULT now(),
    status text NOT NULL,
    mask_year integer,
    activated_at timestamptz,
    validation_result jsonb NOT NULL DEFAULT '{}'::jsonb,
    original_file_path text NOT NULL
);

ALTER TABLE mask_batch ADD COLUMN IF NOT EXISTS mask_year integer;
ALTER TABLE mask_batch ADD COLUMN IF NOT EXISTS activated_at timestamptz;

CREATE TABLE IF NOT EXISTS mask_date (
    mask_batch_id uuid NOT NULL REFERENCES mask_batch(id) ON DELETE CASCADE,
    mask_date date NOT NULL,
    reason text,
    PRIMARY KEY (mask_batch_id, mask_date)
);

CREATE TABLE IF NOT EXISTS batch_activation (
    id uuid PRIMARY KEY,
    batch_id uuid NOT NULL REFERENCES import_batch(id),
    confirmed_by text NOT NULL,
    confirmed_at timestamptz NOT NULL DEFAULT now(),
    replace_start_date date NOT NULL,
    replace_end_date date NOT NULL,
    added_count integer NOT NULL,
    deleted_count integer NOT NULL,
    modified_count integer NOT NULL,
    unchanged_count integer NOT NULL,
    rollback_of uuid REFERENCES batch_activation(id),
    rolled_back_at timestamptz
);

CREATE UNIQUE INDEX IF NOT EXISTS batch_activation_once_idx
    ON batch_activation (batch_id) WHERE rollback_of IS NULL;

CREATE TABLE IF NOT EXISTS report_run (
    id uuid PRIMARY KEY,
    job_id text UNIQUE,
    activation_id uuid REFERENCES batch_activation(id),
    mask_batch_id uuid REFERENCES mask_batch(id),
    parameters jsonb NOT NULL DEFAULT '{}'::jsonb,
    status text NOT NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    started_at timestamptz,
    finished_at timestamptz,
    elapsed_seconds numeric,
    output_files jsonb NOT NULL DEFAULT '{}'::jsonb,
    error_message text
);

CREATE OR REPLACE VIEW current_outage_record AS
SELECT *
FROM outage_record_version
WHERE activated_at IS NOT NULL
  AND superseded_at IS NULL
  AND rolled_back_at IS NULL;
