"use client";

import { useChat } from "@ai-sdk/react";
import type { SpreadsheetAgentUIMessage } from "@repo/core/ai/agent";
import { createAgentUIStream } from "@repo/core/ai/chat";
import {
  type callOptionsSchema,
  messageMetadataSchema,
  models,
} from "@repo/core/ai/schema";
import { type tools, writeTools } from "@repo/core/ai/tools";
import {
  Conversation,
  ConversationContent,
  ConversationScrollButton,
} from "@repo/core/components/ai-elements/conversation";
import { Loader } from "@repo/core/components/ai-elements/loader";
import {
  Message,
  MessageAction,
  MessageActions,
  MessageContent,
  MessageResponse,
} from "@repo/core/components/ai-elements/message";
import {
  PromptInput,
  PromptInputAttachment,
  PromptInputAttachments,
  PromptInputBody,
  PromptInputFooter,
  PromptInputHeader,
  type PromptInputMessage,
  PromptInputSelect,
  PromptInputSelectContent,
  PromptInputSelectItem,
  PromptInputSelectTrigger,
  PromptInputSelectValue,
  PromptInputSubmit,
  PromptInputTextarea,
  PromptInputTools,
} from "@repo/core/components/ai-elements/prompt-input";
import {
  Reasoning,
  ReasoningContent,
  ReasoningTrigger,
} from "@repo/core/components/ai-elements/reasoning";
import {
  Source,
  Sources,
  SourcesContent,
  SourcesTrigger,
} from "@repo/core/components/ai-elements/sources";
import {
  Tool,
  ToolContent,
  ToolHeader,
  ToolInput,
  ToolOutput,
} from "@repo/core/components/ai-elements/tool";
import { ToolApprovalBar } from "@repo/core/components/ai-elements/tool-approval-bar";
import { Anthropic } from "@repo/core/components/icons/anthropic";
import { Button } from "@repo/core/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@repo/core/components/ui/dialog";
import { Input } from "@repo/core/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@repo/core/components/ui/select";
import { useLocalStorage } from "@repo/core/lib/utils";
import type { SpreadsheetService } from "@repo/core/spreadsheet-service";
import {
  type ChatOnToolCallCallback,
  lastAssistantMessageIsCompleteWithToolCalls,
} from "ai";
import {
  ChevronsRightIcon,
  CopyIcon,
  HandIcon,
  PlusIcon,
  RefreshCcwIcon,
  SettingsIcon,
} from "lucide-react";
import { Fragment, useRef, useState } from "react";
import type * as z from "zod";

type CallOptionsSchema = z.infer<typeof callOptionsSchema>;
type Model = CallOptionsSchema["model"];

type DistributiveOmit<T, K extends keyof T> = T extends object
  ? Omit<T, K>
  : never;

const EDIT_MODES = [
  {
    name: "Ask before edits",
    value: "ask",
    icon: HandIcon,
  },
  {
    name: "Accept all edits",
    value: "auto",
    icon: ChevronsRightIcon,
  },
] as const;

function isWriteTool(toolName: keyof typeof tools) {
  return writeTools.includes(toolName as (typeof writeTools)[number]);
}

interface ChatProps {
  spreadsheetService: SpreadsheetService;
  environment: z.infer<typeof callOptionsSchema>["environment"];
}

export function Chat({ spreadsheetService, environment }: ChatProps) {
  const [input, setInput] = useState("");
  const [model, setModel] = useLocalStorage<Model>("model", models[0].value);
  const [anthropicApiKey, setAnthropicApiKey] = useLocalStorage(
    "ANTHROPIC_API_KEY",
    "",
  );
  const [apiKeyInput, setApiKeyInput] = useState("");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [editMode, setEditMode] = useState<"ask" | "auto">(EDIT_MODES[0].value);

  // Update ref synchronously during render to avoid stale closures in transport
  const anthropicApiKeyRef = useRef(anthropicApiKey);
  anthropicApiKeyRef.current = anthropicApiKey;

  const {
    addToolApprovalResponse,
    addToolOutput,
    messages,
    regenerate,
    sendMessage,
    setMessages,
    status,
    stop,
  } = useChat<SpreadsheetAgentUIMessage>({
    messageMetadataSchema,
    transport: {
      async reconnectToStream(_options) {
        throw new Error("Not implemented");
      },
      async sendMessages(options) {
        return createAgentUIStream({
          body: {
            messages: options.messages,
            options: {
              anthropicApiKey: anthropicApiKeyRef.current,
              environment,
              model,
              sheets: await spreadsheetService.getSheets(),
            } satisfies CallOptionsSchema,
          },
        });
      },
    },
    onToolCall: async ({ toolCall }) => {
      if (toolCall.dynamic) return;

      if (isWriteTool(toolCall.toolName)) {
        const input = toolCall.input as Record<string, unknown>;
        const sheetId =
          typeof input.sheetId === "number" ? input.sheetId : undefined;

        if (sheetId !== undefined) {
          let range: string | undefined;

          switch (toolCall.toolName) {
            case "setCellRange":
            case "clearCellRange":
            case "resizeRange":
              range = typeof input.range === "string" ? input.range : undefined;
              break;
            case "copyTo":
              range =
                typeof input.destinationRange === "string"
                  ? input.destinationRange
                  : undefined;
              break;
            case "modifySheetStructure":
            case "modifyObject":
              await spreadsheetService.activateSheet(sheetId);
              break;
          }

          if (range) {
            await spreadsheetService.selectRange({ sheetId, range });
          }
        }

        return;
      }

      await executeTool(toolCall);
    },
    sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls,
  });

  function executeTool(
    toolCall: DistributiveOmit<
      Extract<
        Parameters<
          ChatOnToolCallCallback<SpreadsheetAgentUIMessage>
        >[number]["toolCall"],
        { dynamic?: false }
      >,
      "dynamic"
    >,
  ) {
    const { toolName, toolCallId, input } = toolCall;

    async function run<T>(fn: () => Promise<T>): Promise<void> {
      try {
        const output = await fn();
        addToolOutput({
          state: "output-available",
          tool: toolName,
          toolCallId,
          output,
        });
      } catch (err) {
        const errorText =
          err instanceof Error ? err.message : "An unknown error occurred";
        console.error(`Tool ${toolName} failed:`, err);
        addToolOutput({
          state: "output-error",
          tool: toolName,
          toolCallId,
          errorText,
        });
      } finally {
        if (isWriteTool(toolName)) {
          await spreadsheetService.clearSelection();
        }
      }
    }

    switch (toolName) {
      case "getCellRanges":
        return run(() => spreadsheetService.getCellRanges(input));
      case "searchData":
        return run(() => spreadsheetService.searchData(input));
      case "setCellRange":
        return run(() => spreadsheetService.setCellRange(input));
      case "modifySheetStructure":
        return run(() => spreadsheetService.modifySheetStructure(input));
      case "modifyWorkbookStructure":
        return run(() => spreadsheetService.modifyWorkbookStructure(input));
      case "copyTo":
        return run(() => spreadsheetService.copyTo(input));
      case "getAllObjects":
        return run(() => spreadsheetService.getAllObjects(input));
      case "modifyObject":
        return run(() => spreadsheetService.modifyObject(input));
      case "resizeRange":
        return run(() => spreadsheetService.resizeRange(input));
      case "clearCellRange":
        return run(() => spreadsheetService.clearCellRange(input));
      default:
        console.warn(`Unhandled tool: ${toolName}`);
    }
  }

  function handleSubmit(message: PromptInputMessage) {
    const hasText = Boolean(message.text);
    const hasAttachments = Boolean(message.files?.length);
    if (!(hasText || hasAttachments)) {
      return;
    }

    sendMessage({
      text: message.text || "Sent with attachments",
      files: message.files,
    });

    setInput("");
  }

  return (
    <div className="relative mx-auto size-full h-full">
      <div className="flex h-full flex-col">
        <header className="flex items-center justify-between border-b px-4 py-2">
          <div className="flex items-center gap-2">
            <Button
              variant="ghost"
              size="icon"
              onClick={async () => {
                await stop();
                setMessages([]);
                setInput("");
                setEditMode(EDIT_MODES[0].value);
              }}
            >
              <PlusIcon className="size-4" />
            </Button>
            <Select
              value={model}
              onValueChange={(val) => setModel(val as Model)}
            >
              <SelectTrigger className="border-none shadow-none hover:bg-accent hover:text-accent-foreground">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {models.map((m) => (
                  <SelectItem key={m.value} value={m.value}>
                    <Anthropic className="size-4 fill-[#D97757]" />
                    {m.name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <Dialog
            open={settingsOpen || !anthropicApiKey}
            onOpenChange={(open) => {
              if (open) setApiKeyInput(anthropicApiKey);
              setSettingsOpen(open);
            }}
          >
            <DialogTrigger asChild>
              <Button variant="ghost" size="icon">
                <SettingsIcon className="size-4" />
              </Button>
            </DialogTrigger>
            <DialogContent>
              <DialogHeader>
                <DialogTitle>Settings</DialogTitle>
                <DialogDescription>
                  Configure your API key to use the chat.
                </DialogDescription>
              </DialogHeader>
              <div className="space-y-2 py-2">
                <div className="space-y-2">
                  <label htmlFor="api-key" className="font-medium text-sm">
                    Anthropic API Key
                  </label>
                  <Input
                    id="api-key"
                    type="password"
                    placeholder="sk-ant-..."
                    value={apiKeyInput}
                    onChange={(e) => setApiKeyInput(e.target.value)}
                  />
                  <p className="text-muted-foreground text-xs">
                    Your API key is stored locally and never sent to our
                    servers.
                  </p>
                </div>
              </div>
              <DialogFooter>
                <Button
                  onClick={() => {
                    setAnthropicApiKey(apiKeyInput);
                    setSettingsOpen(false);
                  }}
                >
                  Save
                </Button>
              </DialogFooter>
            </DialogContent>
          </Dialog>
        </header>

        <Conversation className="h-full">
          <ConversationContent className="overflow-x-hidden p-6 pb-16">
            {messages.map((message) => (
              <div key={message.id}>
                {message.role === "assistant" &&
                  message.parts.filter((part) => part.type === "source-url")
                    .length > 0 && (
                    <Sources>
                      <SourcesTrigger
                        count={
                          message.parts.filter(
                            (part) => part.type === "source-url",
                          ).length
                        }
                      />
                      {message.parts
                        .filter((part) => part.type === "source-url")
                        .map((part, i) => (
                          <SourcesContent key={`${message.id}-${i}`}>
                            <Source
                              key={`${message.id}-${i}`}
                              href={part.url}
                              title={part.url}
                            />
                          </SourcesContent>
                        ))}
                    </Sources>
                  )}
                {message.parts.map((part, partIdx) => {
                  switch (part.type) {
                    case "text":
                      return (
                        <Message
                          key={`${message.id}-${partIdx}`}
                          from={message.role}
                        >
                          <MessageContent>
                            <MessageResponse>{part.text}</MessageResponse>
                          </MessageContent>
                          {message.role === "assistant" &&
                            partIdx === messages.length - 1 && (
                              <MessageActions>
                                <MessageAction
                                  onClick={() => regenerate()}
                                  label="Retry"
                                >
                                  <RefreshCcwIcon className="size-3" />
                                </MessageAction>
                                <MessageAction
                                  onClick={() =>
                                    navigator.clipboard.writeText(part.text)
                                  }
                                  label="Copy"
                                >
                                  <CopyIcon className="size-3" />
                                </MessageAction>
                              </MessageActions>
                            )}
                        </Message>
                      );
                    case "reasoning":
                      return (
                        <Reasoning
                          key={`${message.id}-${partIdx}`}
                          className="w-full"
                          isStreaming={
                            status === "streaming" &&
                            partIdx === message.parts.length - 1 &&
                            message.id === messages.at(-1)?.id
                          }
                        >
                          <ReasoningTrigger />
                          <ReasoningContent>{part.text}</ReasoningContent>
                        </Reasoning>
                      );
                    case "tool-bashCodeExecution":
                    case "tool-codeExecution":
                    case "tool-textEditor":
                    case "tool-webSearch":
                    case "tool-clearCellRange":
                    case "tool-copyTo":
                    case "tool-getAllObjects":
                    case "tool-getCellRanges":
                    case "tool-modifyObject":
                    case "tool-modifySheetStructure":
                    case "tool-modifyWorkbookStructure":
                    case "tool-resizeRange":
                    case "tool-searchData":
                    case "tool-setCellRange": {
                      return (
                        <Fragment
                          key={`${message.id}-${partIdx}-${part.state}`}
                        >
                          <Tool
                            defaultOpen={part.state === "approval-requested"}
                          >
                            <ToolHeader
                              state={part.state}
                              type={part.type}
                              title={
                                part.type === "tool-bashCodeExecution"
                                  ? "Executing Bash Code"
                                  : part.type === "tool-codeExecution"
                                    ? "Executing Code"
                                    : part.type === "tool-textEditor"
                                      ? "Editing Text"
                                      : part.type === "tool-webSearch"
                                        ? `Searching "${part.input?.query}"`
                                        : part.input?.explanation
                              }
                            />
                            <ToolContent>
                              <ToolInput
                                toolName={part.type.replace("tool-", "")}
                                input={part.input}
                              />
                              <ToolOutput
                                toolName={part.type.replace("tool-", "")}
                                state={part.state}
                                output={part.output}
                                errorText={part.errorText}
                              />
                            </ToolContent>
                          </Tool>
                          {(() => {
                            if (part.state !== "approval-requested") return;

                            const toolName = part.type.replace(
                              "tool-",
                              "",
                            ) as keyof typeof tools;

                            const approve = () => {
                              addToolApprovalResponse({
                                id: part.approval.id,
                                approved: true,
                              });
                              executeTool({
                                // biome-ignore lint/suspicious/noExplicitAny: <>
                                input: part.input as any,
                                toolCallId: part.toolCallId,
                                toolName,
                              });
                            };

                            if (editMode === "auto") {
                              approve();
                              return null;
                            }

                            return (
                              <ToolApprovalBar
                                key={`${message.id}-${partIdx}-${part.state}`}
                                explanation={
                                  "explanation" in part.input
                                    ? part.input.explanation
                                    : ""
                                }
                                toolName={toolName}
                                onDecline={() => {
                                  addToolApprovalResponse({
                                    id: part.approval.id,
                                    approved: false,
                                  });
                                }}
                                onApprove={() => {
                                  approve();
                                }}
                                onApproveAll={() => {
                                  setEditMode("auto");
                                  approve();
                                }}
                              />
                            );
                          })()}
                        </Fragment>
                      );
                    }
                    default:
                      return null;
                  }
                })}
              </div>
            ))}
            {status === "submitted" && <Loader className="mr-auto" />}
          </ConversationContent>
          <ConversationScrollButton />
        </Conversation>
        <PromptInput
          className="px-3 **:data-[slot=input-group]:rounded-b-none"
          globalDrop
          multiple
          onSubmit={handleSubmit}
        >
          <PromptInputHeader className="p-0!">
            <PromptInputAttachments>
              {(attachment) => (
                <PromptInputAttachment data={attachment} className="truncate" />
              )}
            </PromptInputAttachments>
          </PromptInputHeader>
          <PromptInputBody>
            <PromptInputTextarea
              autoFocus
              className="min-h-24 text-sm!"
              onChange={(e) => setInput(e.target.value)}
              value={input}
            />
          </PromptInputBody>
          <PromptInputFooter>
            <PromptInputTools>
              <PromptInputSelect
                onValueChange={(value: "ask" | "auto") => {
                  setEditMode(value);
                }}
                value={editMode}
              >
                <PromptInputSelectTrigger>
                  <PromptInputSelectValue />
                </PromptInputSelectTrigger>
                <PromptInputSelectContent>
                  {EDIT_MODES.map((mode) => (
                    <PromptInputSelectItem key={mode.value} value={mode.value}>
                      <mode.icon className="size-4" />
                      {mode.name}
                    </PromptInputSelectItem>
                  ))}
                </PromptInputSelectContent>
              </PromptInputSelect>
            </PromptInputTools>
            <PromptInputSubmit disabled={!input && !status} status={status} />
          </PromptInputFooter>
        </PromptInput>
      </div>
    </div>
  );
}
