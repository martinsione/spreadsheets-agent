"use client";

import { useChat } from "@ai-sdk/react";
import {
  type ChatOnToolCallCallback,
  DefaultChatTransport,
  lastAssistantMessageIsCompleteWithToolCalls,
} from "ai";
import { CopyIcon, RefreshCcwIcon } from "lucide-react";
import { useState } from "react";
import {
  Confirmation,
  ConfirmationAction,
  ConfirmationActions,
  ConfirmationTitle,
} from "@/frontend/components/ai-elements/confirmation";
import {
  Conversation,
  ConversationContent,
  ConversationScrollButton,
} from "@/frontend/components/ai-elements/conversation";
import { Loader } from "@/frontend/components/ai-elements/loader";
import {
  Message,
  MessageAction,
  MessageActions,
  MessageContent,
  MessageResponse,
} from "@/frontend/components/ai-elements/message";
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
} from "@/frontend/components/ai-elements/prompt-input";
import {
  Reasoning,
  ReasoningContent,
  ReasoningTrigger,
} from "@/frontend/components/ai-elements/reasoning";
import {
  Source,
  Sources,
  SourcesContent,
  SourcesTrigger,
} from "@/frontend/components/ai-elements/sources";
import {
  Tool,
  ToolContent,
  ToolHeader,
  ToolInput,
  ToolOutput,
} from "@/frontend/components/ai-elements/tool";
import type { SpreadsheetAgentUIMessage } from "@/server/ai/agent";
import { writeTools } from "@/server/ai/tools";
import * as spreadsheetService from "@/spreadsheet-service/excel";

const MODELS = [
  {
    name: "Claude Opus 4.5",
    value: "claude-opus-4-5",
  },
  {
    name: "Claude Sonnet 4.5",
    value: "claude-sonnet-4-5",
  },
] as const;

const EDIT_MODES = [
  {
    name: "Ask before edits",
    value: "ask",
  },
  {
    name: "Accept all edits",
    value: "auto",
  },
] as const;

export default function Chat() {
  const [input, setInput] = useState("");
  const [model] = useState<string>(MODELS[0].value);
  const [editMode, setEditMode] = useState<"ask" | "auto">(EDIT_MODES[0].value);

  const {
    messages,
    sendMessage,
    status,
    regenerate,
    addToolOutput,
    addToolApprovalResponse,
  } = useChat<SpreadsheetAgentUIMessage>({
    transport: new DefaultChatTransport({
      api: "/api/chat",
      prepareSendMessagesRequest: async ({ id, messages }) => ({
        body: {
          id,
          messages,
          model,
          sheets: await spreadsheetService.getSheets(),
        },
      }),
    }),
    onToolCall: async ({ toolCall }) => {
      if (toolCall.dynamic) return;

      const isWriteTool = writeTools.includes(
        toolCall.toolName as (typeof writeTools)[number],
      );

      if (isWriteTool && editMode === "auto") {
        addToolApprovalResponse({ id: toolCall.toolCallId, approved: true });
      }

      await executeTool(toolCall);
    },
    sendAutomaticallyWhen: lastAssistantMessageIsCompleteWithToolCalls,
  });

  function executeTool(
    toolCall: Parameters<
      ChatOnToolCallCallback<SpreadsheetAgentUIMessage>
    >[number]["toolCall"],
  ) {
    if (toolCall.dynamic) return;
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
    <div className="relative mx-auto size-full h-screen max-w-4xl">
      <div className="flex h-full flex-col">
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
                    case "tool-setCellRange":
                      return (
                        <Tool
                          key={`${message.id}-${partIdx}-${part.state}`}
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
                                  : part.type === "tool-webSearch"
                                    ? "Searching the Web"
                                    : part.input?.explanation
                            }
                          />
                          <ToolContent>
                            <ToolInput
                              toolName={part.type.replace("tool-", "")}
                              input={part.input}
                            />
                            <ToolOutput
                              state={part.state}
                              output={part.output}
                              errorText={part.errorText}
                            />
                            <Confirmation
                              state={part.state}
                              approval={part.approval}
                            >
                              <ConfirmationTitle>
                                Allow this action?
                              </ConfirmationTitle>
                              <ConfirmationActions>
                                <ConfirmationAction
                                  variant="outline"
                                  onClick={() =>
                                    addToolApprovalResponse({
                                      // biome-ignore lint/style/noNonNullAssertion: <>
                                      id: part.approval!.id,
                                      approved: false,
                                    })
                                  }
                                >
                                  Deny
                                </ConfirmationAction>
                                <ConfirmationAction
                                  onClick={() => {
                                    const toolName = part.type.replace(
                                      "tool-",
                                      "",
                                    ) as (typeof writeTools)[number];

                                    if (!writeTools.includes(toolName)) {
                                      return;
                                    }

                                    addToolApprovalResponse({
                                      // biome-ignore lint/style/noNonNullAssertion: <>
                                      id: part.approval!.id,
                                      approved: true,
                                    });

                                    executeTool({
                                      // biome-ignore lint/suspicious/noExplicitAny: not worth dealing with this now
                                      input: part.input as any,
                                      toolCallId: part.toolCallId,
                                      toolName,
                                      dynamic: false,
                                    });
                                  }}
                                >
                                  Approve
                                </ConfirmationAction>
                              </ConfirmationActions>
                            </Confirmation>
                          </ToolContent>
                        </Tool>
                      );
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
          autoFocus
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
