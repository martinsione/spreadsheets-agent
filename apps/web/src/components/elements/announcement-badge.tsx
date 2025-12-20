import { clsx } from "clsx/lite";
import type { ComponentProps, ReactNode } from "react";
import { ChevronIcon } from "../icons/chevron-icon";

export function AnnouncementBadge({
  text,
  href,
  cta = "Learn more",
  variant = "normal",
  className,
  ...props
}: {
  text: ReactNode;
  href: string;
  cta?: ReactNode;
  variant?: "normal" | "overlay";
} & Omit<ComponentProps<"a">, "href" | "children">) {
  return (
    <a
      href={href}
      {...props}
      data-variant={variant}
      className={clsx(
        "group relative inline-flex max-w-full gap-x-3 overflow-hidden rounded-md px-3.5 py-2 text-sm/6 max-sm:flex-col sm:items-center sm:rounded-full sm:px-3 sm:py-0.5",
        variant === "normal" &&
          "bg-olive-950/5 text-olive-950 hover:bg-olive-950/10 dark:inset-ring-1 dark:inset-ring-white/5 dark:bg-white/5 dark:text-white dark:hover:bg-white/10",
        variant === "overlay" &&
          "bg-olive-950/15 text-white hover:bg-olive-950/20 dark:bg-olive-950/20 dark:hover:bg-olive-950/25",
        className,
      )}
    >
      <span className="text-pretty sm:truncate">{text}</span>
      <span
        className={clsx(
          "h-3 w-px max-sm:hidden",
          variant === "normal" && "bg-olive-950/20 dark:bg-white/10",
          variant === "overlay" && "bg-white/20",
        )}
      />
      <span
        className={clsx(
          "inline-flex shrink-0 items-center gap-2 font-semibold",
          variant === "normal" && "text-olive-950 dark:text-white",
        )}
      >
        {cta} <ChevronIcon className="shrink-0" />
      </span>
    </a>
  );
}
