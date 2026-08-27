import { withAdmin } from "@/lib/auth/route";
import { CSV_TEMPLATE_HEADERS, CSV_TEMPLATE_ROWS } from "@/lib/integrations/property/csv";

export const dynamic = "force-dynamic";

/** Blank template so an admin can see exactly which columns are expected. */
export const GET = withAdmin(async () => {
  const csv = [CSV_TEMPLATE_HEADERS, ...CSV_TEMPLATE_ROWS].map((row) => row.join(",")).join("\n");
  return new Response(`${csv}\n`, {
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Content-Disposition": 'attachment; filename="property-import-template.csv"',
    },
  });
});
