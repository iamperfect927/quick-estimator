import { google } from '@ai-sdk/google';
import { streamText, tool, zodSchema, convertToModelMessages } from 'ai';
import { z } from 'zod';

export async function POST(req: Request) {
  const { messages } = await req.json();

  const modelMessages = await convertToModelMessages(messages);

  const result = streamText({
    model: google('gemini-1.5-pro'),
    messages: modelMessages,
    system: `You are an expert Solar Installation Estimator Agent. 
    Your job is to analyze the user's project request and prepare a structured dataset for calculations.
    Always execute the 'calculateEstimate' tool once you have gathered material modifications, labor rates, or custom overhead margins.`,
    tools: {
      calculateEstimate: tool({
        description: 'Compiles extracted data to compute final costs and flag for Excel creation.',
        inputSchema: zodSchema(z.object({
          materials: z.array(z.object({
            name: z.string(),
            unitPrice: z.number(),
            quantity: z.number(),
          })),
          labor: z.array(z.object({
            description: z.string(),
            hours: z.number(),
            hourlyRate: z.number(),
          })),
          marginPercentage: z.number().default(15),
        })),
        execute: async (args) => {
          // Simply return data to frontend client context to hold in state
          return { success: true, previewData: args };
        },
      }),
    },
  });

  return result.toUIMessageStreamResponse();
}