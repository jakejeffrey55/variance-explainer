import { z } from "zod";

export const propertyInputSchema = z.object({
  externalId: z.string().trim().max(64).optional().nullable(),
  name: z.string().trim().min(2, "Give the property a name.").max(160),
  addressLine1: z.string().trim().min(3, "Street address is required.").max(200),
  addressLine2: z.string().trim().max(200).optional().nullable(),
  city: z.string().trim().min(1, "City is required.").max(120),
  state: z.string().trim().min(2, "State is required.").max(32),
  postalCode: z.string().trim().min(3, "Postal code is required.").max(16),
  latitude: z
    .number({ error: "Latitude is required — it drives vendor radius matching." })
    .min(-90)
    .max(90),
  longitude: z
    .number({ error: "Longitude is required — it drives vendor radius matching." })
    .min(-180)
    .max(180),
  unitCount: z.number().int().min(0).max(100000).optional().nullable(),
  propertyManagerName: z.string().trim().max(160).optional().nullable(),
  propertyManagerEmail: z.string().trim().email("Enter a valid email.").optional().nullable().or(z.literal("")),
  propertyManagerPhone: z.string().trim().max(32).optional().nullable(),
  isActive: z.boolean().default(true),
});

export const propertyUpdateSchema = propertyInputSchema.partial();

export type PropertyInput = z.infer<typeof propertyInputSchema>;
