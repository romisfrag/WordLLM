import { ChatOpenAI } from "@langchain/openai";


// Function to initialize the model with custom configuration
export function initializeModel(baseURL: string, openAIApiKey: string) {
    return new ChatOpenAI({
        openAIApiKey: openAIApiKey,
        configuration: {
            baseURL: baseURL,
            defaultHeaders: {
                "HTTP-Referer": "https://localhost:3000",
                "X-Title": "WordLLM",
            },
        },
        modelName: "qwen/qwen-2.5-7b-instruct:free",
        temperature: 0.7,
    });
}

// Function to get response from the LLM
export async function getLLMResponse(prompt: string, model: ChatOpenAI): Promise<string> {
    try {
        const response = await model.invoke(prompt);
        return response.content as string;
    } catch (error: any) {
        console.error("Error getting LLM response:", error);
        
        // Extract error details from the response if available
        let errorMessage = "Sorry, there was an error processing your request.";
        if (error.response) {
            try {
                const errorData = await error.response.json();
                errorMessage = `API Error: ${errorData.error?.message || JSON.stringify(errorData)}`;
            } catch (e) {
                errorMessage = `API Error: ${error.response.status} ${error.response.statusText}`;
            }
        } else if (error.message) {
            errorMessage = `Error: ${error.message}`;
        }
        
        return errorMessage;
    }
} 