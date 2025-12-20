import { clsx } from "clsx/lite";
import type { ComponentProps } from "react";

export function Subheading({
  children,
  className,
  ...props
}: ComponentProps<"h2">) {
  return (
    <h2
      className={clsx(
        "text-pretty font-display text-[2rem]/10 text-olive-950 tracking-tight sm:text-5xl/14 dark:text-white",
        className,
      )}
      {...props}
    >
      {children}
    </h2>
  );
}
