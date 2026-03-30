/*
  Warnings:

  - You are about to drop the column `isPersonal` on the `ConversationRef` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "ConversationRef" DROP COLUMN "isPersonal";

-- AlterTable
ALTER TABLE "Employee" ADD COLUMN     "carry_forward" DOUBLE PRECISION NOT NULL DEFAULT 0,
ALTER COLUMN "leave_balance" SET DEFAULT 0;

-- AlterTable
ALTER TABLE "LeaveRequest" ADD COLUMN     "lop_days" DOUBLE PRECISION NOT NULL DEFAULT 0;

-- AlterTable
ALTER TABLE "PendingRequest" ADD COLUMN     "lop_days" DOUBLE PRECISION NOT NULL DEFAULT 0;

-- CreateTable
CREATE TABLE "MonthlySummary" (
    "id" SERIAL NOT NULL,
    "month" TEXT NOT NULL,
    "employee" TEXT NOT NULL,
    "opening" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "available" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "leaves" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "wfh" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "lop" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "closing" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "pending" DOUBLE PRECISION NOT NULL DEFAULT 0,
    "created_at" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "MonthlySummary_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "MonthlySummary_month_employee_key" ON "MonthlySummary"("month", "employee");

-- AddForeignKey
ALTER TABLE "MonthlySummary" ADD CONSTRAINT "MonthlySummary_employee_fkey" FOREIGN KEY ("employee") REFERENCES "Employee"("name") ON DELETE RESTRICT ON UPDATE CASCADE;
