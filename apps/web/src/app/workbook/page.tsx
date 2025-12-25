"use client";

import dynamic from "next/dynamic";

const WorkbookContent = dynamic(
  () => import("./workbook-content").then((mod) => mod.WorkbookContent),
  {
    ssr: false,
    loading: () => (
      <div className="flex h-full w-full animate-pulse items-center justify-center text-muted-foreground duration-300">
        Loading...
      </div>
    ),
  },
);

export default function WorkbookPage() {
  return <WorkbookContent />;
}
