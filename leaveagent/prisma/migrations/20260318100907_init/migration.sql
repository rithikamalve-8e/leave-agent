-- CreateTable
CREATE TABLE "Employee" (
    "id" SERIAL NOT NULL,
    "name" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "role" TEXT NOT NULL,
    "bot_role" TEXT NOT NULL DEFAULT 'employee',
    "manager" TEXT,
    "manager_email" TEXT,
    "manager_teams_id" TEXT,
    "teamlead" TEXT,
    "teamlead_email" TEXT,
    "teamlead_teams_id" TEXT,
    "teams_id" TEXT,
    "leave_balance" DOUBLE PRECISION NOT NULL DEFAULT 22,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Employee_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "LeaveRequest" (
    "id" SERIAL NOT NULL,
    "employee" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "type" TEXT NOT NULL,
    "date" TEXT NOT NULL,
    "end_date" TEXT,
    "duration" TEXT NOT NULL,
    "days_count" DOUBLE PRECISION NOT NULL DEFAULT 1,
    "reason" TEXT,
    "rejection_reason" TEXT,
    "status" TEXT NOT NULL DEFAULT 'Pending',
    "approved_by" TEXT,
    "deleted_by" TEXT,
    "deleted_at" TIMESTAMP(3),
    "requested_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "LeaveRequest_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ConversationRef" (
    "id" SERIAL NOT NULL,
    "userId" TEXT NOT NULL,
    "userName" TEXT NOT NULL,
    "conversationId" TEXT NOT NULL,
    "serviceUrl" TEXT NOT NULL,
    "tenantId" TEXT,
    "botId" TEXT NOT NULL,
    "isPersonal" BOOLEAN NOT NULL DEFAULT false,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "ConversationRef_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Holiday" (
    "id" SERIAL NOT NULL,
    "date" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "added_by" TEXT NOT NULL,
    "added_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "Holiday_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "AuditLog" (
    "id" SERIAL NOT NULL,
    "timestamp" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "hr_name" TEXT NOT NULL,
    "action" TEXT NOT NULL,
    "target_employee" TEXT,
    "details" TEXT NOT NULL,

    CONSTRAINT "AuditLog_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PendingRequest" (
    "id" SERIAL NOT NULL,
    "userId" TEXT NOT NULL,
    "userName" TEXT NOT NULL,
    "intent" TEXT NOT NULL,
    "date" TEXT NOT NULL,
    "end_date" TEXT,
    "duration" TEXT NOT NULL,
    "days_count" DOUBLE PRECISION NOT NULL,
    "reason" TEXT,
    "balance_json" TEXT NOT NULL,
    "history_json" TEXT NOT NULL,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PendingRequest_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "Employee_name_key" ON "Employee"("name");

-- CreateIndex
CREATE UNIQUE INDEX "Employee_email_key" ON "Employee"("email");

-- CreateIndex
CREATE UNIQUE INDEX "Employee_teams_id_key" ON "Employee"("teams_id");

-- CreateIndex
CREATE UNIQUE INDEX "ConversationRef_userId_key" ON "ConversationRef"("userId");

-- CreateIndex
CREATE UNIQUE INDEX "ConversationRef_userName_key" ON "ConversationRef"("userName");

-- CreateIndex
CREATE UNIQUE INDEX "Holiday_date_key" ON "Holiday"("date");

-- CreateIndex
CREATE UNIQUE INDEX "PendingRequest_userId_key" ON "PendingRequest"("userId");

-- AddForeignKey
ALTER TABLE "LeaveRequest" ADD CONSTRAINT "LeaveRequest_employee_fkey" FOREIGN KEY ("employee") REFERENCES "Employee"("name") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ConversationRef" ADD CONSTRAINT "ConversationRef_userName_fkey" FOREIGN KEY ("userName") REFERENCES "Employee"("name") ON DELETE RESTRICT ON UPDATE CASCADE;
