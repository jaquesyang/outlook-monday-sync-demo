-- CreateTable
CREATE TABLE "users" (
    "id" UUID NOT NULL,
    "ms_tenant_id" TEXT NOT NULL,
    "ms_user_id" TEXT NOT NULL,
    "ms_user_email" TEXT NOT NULL,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "users_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ms_accounts" (
    "user_id" UUID NOT NULL,
    "access_token_enc" BYTEA NOT NULL,
    "refresh_token_enc" BYTEA NOT NULL,
    "expires_at" TIMESTAMP(3) NOT NULL,
    "graph_subscription_id" TEXT,
    "graph_subscription_expires_at" TIMESTAMP(3),
    "selected_calendar_id" TEXT,

    CONSTRAINT "ms_accounts_pkey" PRIMARY KEY ("user_id")
);

-- CreateTable
CREATE TABLE "monday_accounts" (
    "user_id" UUID NOT NULL,
    "monday_user_id" BIGINT NOT NULL,
    "monday_account_id" BIGINT NOT NULL,
    "access_token_enc" BYTEA NOT NULL,
    "refresh_token_enc" BYTEA NOT NULL,
    "expires_at" TIMESTAMP(3) NOT NULL,
    "default_board_id" BIGINT,
    "default_group_id" TEXT,

    CONSTRAINT "monday_accounts_pkey" PRIMARY KEY ("user_id")
);

-- CreateTable
CREATE TABLE "board_subscriptions" (
    "id" UUID NOT NULL,
    "user_id" UUID NOT NULL,
    "monday_board_id" BIGINT NOT NULL,
    "date_column_id" TEXT NOT NULL,
    "date_column_kind" TEXT NOT NULL,
    "monday_webhook_id" BIGINT,
    "active" BOOLEAN NOT NULL DEFAULT true,

    CONSTRAINT "board_subscriptions_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "event_mappings" (
    "id" UUID NOT NULL,
    "user_id" UUID NOT NULL,
    "monday_item_id" BIGINT NOT NULL,
    "monday_board_id" BIGINT NOT NULL,
    "graph_event_id" TEXT NOT NULL,
    "graph_calendar_id" TEXT NOT NULL,
    "monday_etag" TEXT,
    "graph_etag" TEXT,
    "origin" TEXT NOT NULL,
    "last_synced_at" TIMESTAMP(3),
    "deleted_at" TIMESTAMP(3),

    CONSTRAINT "event_mappings_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "sync_log" (
    "id" UUID NOT NULL,
    "user_id" UUID NOT NULL,
    "direction" TEXT NOT NULL,
    "mapping_id" UUID,
    "action" TEXT NOT NULL,
    "message" TEXT,
    "occurred_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "sync_log_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "users_ms_tenant_id_ms_user_id_key" ON "users"("ms_tenant_id", "ms_user_id");

-- CreateIndex
CREATE UNIQUE INDEX "board_subscriptions_user_id_monday_board_id_key" ON "board_subscriptions"("user_id", "monday_board_id");

-- CreateIndex
CREATE UNIQUE INDEX "event_mappings_user_id_monday_item_id_key" ON "event_mappings"("user_id", "monday_item_id");

-- CreateIndex
CREATE UNIQUE INDEX "event_mappings_user_id_graph_event_id_key" ON "event_mappings"("user_id", "graph_event_id");

-- CreateIndex
CREATE INDEX "sync_log_user_id_occurred_at_idx" ON "sync_log"("user_id", "occurred_at");

-- AddForeignKey
ALTER TABLE "ms_accounts" ADD CONSTRAINT "ms_accounts_user_id_fkey" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "monday_accounts" ADD CONSTRAINT "monday_accounts_user_id_fkey" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "board_subscriptions" ADD CONSTRAINT "board_subscriptions_user_id_fkey" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "event_mappings" ADD CONSTRAINT "event_mappings_user_id_fkey" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "sync_log" ADD CONSTRAINT "sync_log_user_id_fkey" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE CASCADE ON UPDATE CASCADE;
