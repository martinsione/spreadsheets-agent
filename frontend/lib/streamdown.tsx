"use client";

/**
 * Lightweight Streamdown implementation
 * Uses the same deps as streamdown but avoids importing from the main package
 * which bundles all shiki languages and mermaid
 */

import type { Element, Nodes } from "hast";
import { toJsxRuntime } from "hast-util-to-jsx-runtime";
import { Lexer } from "marked";
import type { ComponentType, CSSProperties, JSX, ReactElement } from "react";
import {
  createContext,
  memo,
  useContext,
  useEffect,
  useId,
  useMemo,
  useState,
  useTransition,
} from "react";
import { Fragment, jsx, jsxs } from "react/jsx-runtime";
import remarkGfm from "remark-gfm";
import remarkParse from "remark-parse";
import type { Options as RemarkRehypeOptions } from "remark-rehype";
import remarkRehype from "remark-rehype";
import remend from "remend";

import type { PluggableList } from "unified";
import { unified } from "unified";
import { cn } from "./utils";

// ============================================================================
// Types
// ============================================================================

export type ExtraProps = { node?: Element | undefined };

export type Components = {
  [Key in keyof JSX.IntrinsicElements]?:
    | ComponentType<JSX.IntrinsicElements[Key] & ExtraProps>
    | keyof JSX.IntrinsicElements;
};

export type MarkdownOptions = {
  children?: string;
  components?: Components;
  rehypePlugins?: PluggableList;
  remarkPlugins?: PluggableList;
  remarkRehypeOptions?: Readonly<RemarkRehypeOptions>;
};

export type StreamdownProps = MarkdownOptions & {
  mode?: "static" | "streaming";
  parseIncompleteMarkdown?: boolean;
  className?: string;
  isAnimating?: boolean;
  caret?: "block" | "circle";
};

type BlockProps = MarkdownOptions & {
  content: string;
  index: number;
};

// ============================================================================
// Context
// ============================================================================

export type StreamdownContextType = {
  isAnimating: boolean;
  mode: "static" | "streaming";
};

export const StreamdownContext = createContext<StreamdownContextType>({
  isAnimating: false,
  mode: "streaming",
});

export const useStreamdownContext = () => useContext(StreamdownContext);

// ============================================================================
// Parse Markdown Into Blocks (from marked Lexer)
// ============================================================================

const footnoteRefPattern = /\[\^[^\]\s]{1,200}\](?!:)/;
const footnoteDefPattern = /\[\^[^\]\s]{1,200}\]:/;

export const parseMarkdownIntoBlocks = (markdown: string): string[] => {
  if (footnoteRefPattern.test(markdown) || footnoteDefPattern.test(markdown)) {
    return [markdown];
  }

  const tokens = Lexer.lex(markdown, { gfm: true });
  return tokens.map((t) => t.raw);
};

// ============================================================================
// Markdown Renderer (unified pipeline)
// ============================================================================

const EMPTY_PLUGINS: PluggableList = [];
const DEFAULT_OPTS = { allowDangerousHtml: true };

// biome-ignore lint/suspicious/noExplicitAny: Processor types are complex
const processorCache = new Map<string, any>();

const getProcessor = (opts: Readonly<MarkdownOptions>) => {
  const key = JSON.stringify([
    opts.remarkPlugins?.map((p) =>
      Array.isArray(p) ? p[0]?.name : (p as { name?: string })?.name,
    ),
    opts.rehypePlugins?.map((p) =>
      Array.isArray(p) ? p[0]?.name : (p as { name?: string })?.name,
    ),
  ]);

  const cached = processorCache.get(key);
  if (cached) return cached;

  const proc = unified()
    .use(remarkParse)
    .use(opts.remarkPlugins || EMPTY_PLUGINS)
    .use(
      remarkRehype,
      opts.remarkRehypeOptions
        ? { ...DEFAULT_OPTS, ...opts.remarkRehypeOptions }
        : DEFAULT_OPTS,
    )
    .use(opts.rehypePlugins || EMPTY_PLUGINS);

  processorCache.set(key, proc);
  return proc;
};

export const Markdown = (opts: Readonly<MarkdownOptions>): ReactElement => {
  const proc = getProcessor(opts);
  const content = opts.children || "";
  const tree = proc.runSync(proc.parse(content), content) as Nodes;
  return toJsxRuntime(tree, {
    Fragment,
    components: opts.components,
    ignoreInvalidStyle: true,
    jsx,
    jsxs,
    passKeys: true,
    passNode: true,
  });
};

// ============================================================================
// Default Components (simple, no shiki/mermaid)
// ============================================================================

const LANG_REGEX = /language-([^\s]+)/;

const defaultComponents: Components = {
  ol: ({ className, children, ...p }) => (
    <ol
      className={cn("list-inside list-decimal in-[li]:pl-6", className)}
      {...p}
    >
      {children}
    </ol>
  ),
  ul: ({ className, children, ...p }) => (
    <ul className={cn("list-inside list-disc in-[li]:pl-6", className)} {...p}>
      {children}
    </ul>
  ),
  li: ({ className, children, ...p }) => (
    <li className={cn("py-1 [&>p]:inline", className)} {...p}>
      {children}
    </li>
  ),
  hr: ({ className, ...p }) => (
    <hr className={cn("my-6 border-border", className)} {...p} />
  ),
  strong: ({ className, children, ...p }) => (
    <span className={cn("font-semibold", className)} {...p}>
      {children}
    </span>
  ),
  a: ({ className, children, href, ...p }) => (
    <a
      className={cn("font-medium text-primary underline", className)}
      href={href}
      rel="noreferrer"
      target="_blank"
      {...p}
    >
      {children}
    </a>
  ),
  h1: ({ className, children, ...p }) => (
    <h1 className={cn("mt-6 mb-2 font-semibold text-3xl", className)} {...p}>
      {children}
    </h1>
  ),
  h2: ({ className, children, ...p }) => (
    <h2 className={cn("mt-6 mb-2 font-semibold text-2xl", className)} {...p}>
      {children}
    </h2>
  ),
  h3: ({ className, children, ...p }) => (
    <h3 className={cn("mt-6 mb-2 font-semibold text-xl", className)} {...p}>
      {children}
    </h3>
  ),
  h4: ({ className, children, ...p }) => (
    <h4 className={cn("mt-6 mb-2 font-semibold text-lg", className)} {...p}>
      {children}
    </h4>
  ),
  h5: ({ className, children, ...p }) => (
    <h5 className={cn("mt-6 mb-2 font-semibold text-base", className)} {...p}>
      {children}
    </h5>
  ),
  h6: ({ className, children, ...p }) => (
    <h6 className={cn("mt-6 mb-2 font-semibold text-sm", className)} {...p}>
      {children}
    </h6>
  ),
  blockquote: ({ className, children, ...p }) => (
    <blockquote
      className={cn(
        "my-4 border-muted-foreground/30 border-l-4 pl-4 text-muted-foreground italic",
        className,
      )}
      {...p}
    >
      {children}
    </blockquote>
  ),
  table: ({ className, children, ...p }) => (
    <div className="my-4 w-full overflow-auto">
      <table
        className={cn("w-full border-collapse rounded-lg border", className)}
        {...p}
      >
        {children}
      </table>
    </div>
  ),
  thead: ({ className, children, ...p }) => (
    <thead className={cn("bg-muted/80", className)} {...p}>
      {children}
    </thead>
  ),
  tbody: ({ className, children, ...p }) => (
    <tbody
      className={cn("divide-y divide-border bg-muted/40", className)}
      {...p}
    >
      {children}
    </tbody>
  ),
  tr: ({ className, children, ...p }) => (
    <tr className={cn("border-border border-b", className)} {...p}>
      {children}
    </tr>
  ),
  th: ({ className, children, ...p }) => (
    <th
      className={cn(
        "whitespace-nowrap px-4 py-2 text-left font-semibold text-sm",
        className,
      )}
      {...p}
    >
      {children}
    </th>
  ),
  td: ({ className, children, ...p }) => (
    <td className={cn("px-4 py-2 text-sm", className)} {...p}>
      {children}
    </td>
  ),
  pre: ({ children }) => <>{children}</>,
  code: ({ node, className, children, ...p }) => {
    const inline = node?.position?.start.line === node?.position?.end.line;
    if (inline) {
      return (
        <code
          className={cn(
            "rounded bg-muted px-1.5 py-0.5 font-mono text-sm",
            className,
          )}
          {...p}
        >
          {children}
        </code>
      );
    }
    const lang = className?.match(LANG_REGEX)?.[1] ?? "";
    return (
      <div className="group relative my-4 overflow-hidden rounded-xl border bg-muted/30">
        {lang && (
          <div className="border-b bg-muted/50 px-4 py-2 font-mono text-muted-foreground text-xs">
            {lang}
          </div>
        )}
        <pre className="overflow-x-auto p-4">
          <code className="font-mono text-sm" {...p}>
            {children}
          </code>
        </pre>
      </div>
    );
  },
};

// ============================================================================
// Block Component
// ============================================================================

const defaultRemarkPlugins: PluggableList = [[remarkGfm, {}]];

const Block = memo(
  ({ content, ...props }: BlockProps) => (
    <Markdown {...props}>{content}</Markdown>
  ),
  (prev, next) => prev.content === next.content && prev.index === next.index,
);
Block.displayName = "Block";

// ============================================================================
// Streamdown Component
// ============================================================================

const carets = { block: " ▋", circle: " ●" };

export const Streamdown = memo(
  ({
    children,
    mode = "streaming",
    parseIncompleteMarkdown = true,
    components,
    rehypePlugins = EMPTY_PLUGINS,
    remarkPlugins = defaultRemarkPlugins,
    className,
    isAnimating = false,
    caret,
    ...props
  }: StreamdownProps) => {
    const id = useId();
    const [, startTransition] = useTransition();
    const [displayBlocks, setDisplayBlocks] = useState<string[]>([]);

    const processed = useMemo(() => {
      if (typeof children !== "string") return "";
      return mode === "streaming" && parseIncompleteMarkdown
        ? remend(children)
        : children;
    }, [children, mode, parseIncompleteMarkdown]);

    const blocks = useMemo(
      () => parseMarkdownIntoBlocks(processed),
      [processed],
    );

    useEffect(() => {
      if (mode === "streaming") {
        startTransition(() => setDisplayBlocks(blocks));
      } else {
        setDisplayBlocks(blocks);
      }
    }, [blocks, mode]);

    const toRender = mode === "streaming" ? displayBlocks : blocks;
    const keys = useMemo(
      () => toRender.map((_, i) => `${id}-${i}`),
      [toRender, id],
    );

    const ctx = useMemo<StreamdownContextType>(
      () => ({ isAnimating, mode }),
      [isAnimating, mode],
    );
    const merged = useMemo(
      () => ({ ...defaultComponents, ...components }),
      [components],
    );

    const style = useMemo(
      () =>
        caret && isAnimating
          ? ({ "--streamdown-caret": `"${carets[caret]}"` } as CSSProperties)
          : undefined,
      [caret, isAnimating],
    );

    if (mode === "static") {
      return (
        <StreamdownContext.Provider value={ctx}>
          <div className={cn("space-y-4 *:first:mt-0 *:last:mb-0", className)}>
            <Markdown
              components={merged}
              rehypePlugins={rehypePlugins}
              remarkPlugins={remarkPlugins}
              {...props}
            >
              {children}
            </Markdown>
          </div>
        </StreamdownContext.Provider>
      );
    }

    return (
      <StreamdownContext.Provider value={ctx}>
        <div
          className={cn(
            "space-y-4 *:first:mt-0 *:last:mb-0",
            caret &&
              "*:last:after:inline *:last:after:align-baseline *:last:after:content-(--streamdown-caret)",
            className,
          )}
          style={style}
        >
          {toRender.map((block, i) => (
            <Block
              key={keys[i]}
              content={block}
              index={i}
              components={merged}
              rehypePlugins={rehypePlugins}
              remarkPlugins={remarkPlugins}
              {...props}
            />
          ))}
        </div>
      </StreamdownContext.Provider>
    );
  },
  (prev, next) =>
    prev.children === next.children &&
    prev.isAnimating === next.isAnimating &&
    prev.mode === next.mode,
);
Streamdown.displayName = "Streamdown";
