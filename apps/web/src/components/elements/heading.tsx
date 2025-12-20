import { clsx } from "clsx/lite";
import type { ComponentProps } from "react";

export function Heading({
  children,
  color = "dark/light",
  className,
  ...props
}: { color?: "dark/light" | "light" } & ComponentProps<"h1">) {
  return (
    <h1
      className={clsx(
        "text-balance font-display text-5xl/12 tracking-tight sm:text-[5rem]/20",
        color === "dark/light" && "text-olive-950 dark:text-white",
        color === "light" && "text-white",
        className,
      )}
      {...props}
    >
      {children}
    </h1>
  );
}
