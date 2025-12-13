import * as z from "zod";

export const Sheet = z.object({
  id: z.int(),
  name: z.string(),
  maxRows: z.int(),
  maxColumns: z.int(),
});
export type Sheet = z.infer<typeof Sheet>;
