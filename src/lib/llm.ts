import { ChatOpenAI } from "@langchain/openai";

// const openRouterApiKey = process.env.OPENROUTER_API_KEY || "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";
const openRouterApiKey = "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";

// Initialize the OpenAI chat model
const model = new ChatOpenAI({
    openAIApiKey: openRouterApiKey,
    configuration: {
        baseURL: "https://openrouter.ai/api/v1", // Important: use OpenRouter's endpoint
        defaultHeaders: {
            "HTTP-Referer": "https://localhost:3000",
            "X-Title": "WordLLM", // optional title for usage logs
        },
    },
    modelName: "qwen/qwen-2.5-7b-instruct:free",
    temperature: 0.7,
});

// Function to get response from the LLM
export async function getLLMResponse(prompt: string): Promise<string> {
    try {
        const response = await model.invoke(prompt);
        return response.content as string;
    } catch (error) {
        console.error("Error getting LLM response:", error);
        return "Sorry, there was an error processing your request.";
    }
} 