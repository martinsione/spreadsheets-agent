import type { ReactNode } from "react";

type MDXComponents = Record<
  string,
  (props: Record<string, unknown>) => ReactNode
>;

export function useMDXComponents(components: MDXComponents): MDXComponents {
  return {
    ...components,
  };
}
