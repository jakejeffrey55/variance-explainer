import type { EmergencyCategory, JobPriority, ServiceCategory } from "@prisma/client";

/**
 * Form-shaped job values (all strings, as HTML inputs produce them).
 *
 * Lives outside the client component so server components can build the
 * initial value — exports from a "use client" module are client references and
 * cannot be called on the server.
 */
export type JobFormValues = {
  id?: string;
  propertyId: string;
  unitNumber: string;
  title: string;
  description: string;
  category: ServiceCategory;
  budgetMin: string;
  budgetMax: string;
  enforceBudgetCap: boolean;
  bidDeadline: string;
  scheduledStart: string;
  dueDate: string;
  priority: JobPriority;
  emergencyCategory: EmergencyCategory | "";
  responseDeadlineMinutes: string;
  inviteOnly: boolean;
};

export const blankJob = (propertyId = ""): JobFormValues => ({
  propertyId,
  unitNumber: "",
  title: "",
  description: "",
  category: "MAKE_READY",
  budgetMin: "",
  budgetMax: "",
  enforceBudgetCap: false,
  bidDeadline: "",
  scheduledStart: "",
  dueDate: "",
  priority: "STANDARD",
  emergencyCategory: "",
  responseDeadlineMinutes: "15",
  inviteOnly: false,
});
