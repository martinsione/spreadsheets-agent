"use client";

import { useChat } from "@ai-sdk/react";
import {
  lastAssistantMessageIsCompleteWithApprovalResponses,
  lastAssistantMessageIsCompleteWithToolCalls,
} from "ai";
import { CopyIcon, RefreshCcwIcon } from "lucide-react";
import { useState } from "react";
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
  PromptInputActionAddAttachments,
  PromptInputActionMenu,
  PromptInputActionMenuContent,
  PromptInputActionMenuTrigger,
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

function lastAssistantMessageIsCompleteWithApprovalResponsesExcludingClientSideTools({
  messages,
}: {
  messages: SpreadsheetAgentUIMessage[];
}) {
  const lastMessage = messages.at(-1);
  if (lastMessage?.role !== "assistant") {
    return false;
  }

  return (
    lastAssistantMessageIsCompleteWithApprovalResponses({ messages }) &&
    !lastMessage.parts.some((part) => writeTools.includes(part.type))
  );
}

export default function Chat() {
  const [input, setInput] = useState("");
  const [model, setModel] = useState<string>(MODELS[0].value);
  const { messages, sendMessage, status, regenerate, addToolOutput } =
    useChat<SpreadsheetAgentUIMessage>({
      sendAutomaticallyWhen: ({ messages }) =>
        lastAssistantMessageIsCompleteWithToolCalls({ messages }) ||
        lastAssistantMessageIsCompleteWithApprovalResponsesExcludingClientSideTools(
          { messages },
        ),
      onToolCall: async ({ toolCall }) => {
        // Must check !toolCall.dynamic to exclude dynamic tool calls from the union.
        // Without this, TypeScript can't narrow the input type because the dynamic
        // case (toolName: string, input: unknown) also matches any toolName check.
        if (toolCall.dynamic) return;

        const { toolCallId, toolName: tool, input } = toolCall;
        const state = "output-available" as const;

        switch (tool) {
          case "getCellRanges": {
            const output = await spreadsheetService.getCellRanges(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "searchData": {
            const output = await spreadsheetService.searchData(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "setCellRange": {
            const output = await spreadsheetService.setCellRange(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "modifySheetStructure": {
            const output = await spreadsheetService.modifySheetStructure(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "modifyWorkbookStructure": {
            const output =
              await spreadsheetService.modifyWorkbookStructure(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "copyTo": {
            const output = await spreadsheetService.copyTo(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "getAllObjects": {
            const output = await spreadsheetService.getAllObjects(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "modifyObject": {
            const output = await spreadsheetService.modifyObject(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "resizeRange": {
            const output = await spreadsheetService.resizeRange(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
          case "clearCellRange": {
            const output = await spreadsheetService.clearCellRange(input);
            addToolOutput({ state, tool, toolCallId, output });
            break;
          }
        }
      },
    });

  const handleSubmit = async (message: PromptInputMessage) => {
    const hasText = Boolean(message.text);
    const hasAttachments = Boolean(message.files?.length);
    if (!(hasText || hasAttachments)) {
      return;
    }

    const sheets = await spreadsheetService.getSheets();

    sendMessage(
      {
        text: message.text || "Sent with attachments",
        files: message.files,
      },
      {
        body: { model, sheets },
      },
    );

    setInput("");
  };

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
                {message.parts.map((part, i) => {
                  switch (part.type) {
                    case "text":
                      return (
                        <Message key={`${message.id}-${i}`} from={message.role}>
                          <MessageContent>
                            <MessageResponse>{part.text}</MessageResponse>
                          </MessageContent>
                          {message.role === "assistant" &&
                            i === messages.length - 1 && (
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
                          key={`${message.id}-${i}`}
                          className="w-full"
                          isStreaming={
                            status === "streaming" &&
                            i === message.parts.length - 1 &&
                            message.id === messages.at(-1)?.id
                          }
                        >
                          <ReasoningTrigger />
                          <ReasoningContent>{part.text}</ReasoningContent>
                        </Reasoning>
                      );
                    case "tool-bashCodeExecution":
                    case "tool-clearCellRange":
                    case "tool-codeExecution":
                    case "tool-copyTo":
                    case "tool-getAllObjects":
                    case "tool-getCellRanges":
                    case "tool-modifyObject":
                    case "tool-modifySheetStructure":
                    case "tool-modifyWorkbookStructure":
                    case "tool-resizeRange":
                    case "tool-searchData":
                    case "tool-setCellRange":
                    case "tool-webSearch":
                      return (
                        <pre
                          key={`${message.id}-${i}`}
                          className="text-muted-foreground text-xs"
                        >
                          {JSON.stringify(part, null, 2)}
                        </pre>
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
          className="px-2 **:data-[slot=input-group]:rounded-b-none"
          globalDrop
          multiple
          onSubmit={handleSubmit}
        >
          <PromptInputHeader>
            <PromptInputAttachments>
              {(attachment) => <PromptInputAttachment data={attachment} />}
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
              <PromptInputActionMenu>
                <PromptInputActionMenuTrigger />
                <PromptInputActionMenuContent>
                  <PromptInputActionAddAttachments />
                </PromptInputActionMenuContent>
              </PromptInputActionMenu>
              <PromptInputSelect
                onValueChange={(value) => {
                  setModel(value);
                }}
                value={model}
              >
                <PromptInputSelectTrigger>
                  <PromptInputSelectValue />
                </PromptInputSelectTrigger>
                <PromptInputSelectContent>
                  {MODELS.map((model) => (
                    <PromptInputSelectItem
                      key={model.value}
                      value={model.value}
                    >
                      {model.name}
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
