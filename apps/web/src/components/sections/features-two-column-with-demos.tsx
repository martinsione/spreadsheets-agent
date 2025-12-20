import { clsx } from "clsx/lite";
import type { ComponentProps, ReactNode } from "react";
import { Section } from "../elements/section";

export function Feature({
  demo,
  headline,
  subheadline,
  cta,
  className,
}: {
  demo: ReactNode;
  headline: ReactNode;
  subheadline: ReactNode;
  cta: ReactNode;
} & Omit<ComponentProps<"div">, "children">) {
  return (
    <div
      className={clsx(
        "rounded-lg bg-olive-950/2.5 p-2 dark:bg-white/5",
        className,
      )}
    >
      <div className="relative overflow-hidden rounded-sm dark:after:absolute dark:after:inset-0 dark:after:rounded-sm dark:after:outline-1 dark:after:outline-white/10 dark:after:-outline-offset-1">
        {demo}
      </div>
      <div className="flex flex-col gap-4 p-6 sm:p-10 lg:p-6">
        <div>
          <h3 className="font-medium text-base/8 text-olive-950 dark:text-white">
            {headline}
          </h3>
          <div className="mt-2 flex flex-col gap-4 text-olive-700 text-sm/7 dark:text-olive-400">
            {subheadline}
          </div>
        </div>
        {cta}
      </div>
    </div>
  );
}

export function FeaturesTwoColumnWithDemos({
  features,
  ...props
}: { features: ReactNode } & Omit<ComponentProps<typeof Section>, "children">) {
  return (
    <Section {...props}>
      <div className="grid grid-cols-1 gap-2 lg:grid-cols-2">{features}</div>
    </Section>
  );
}
