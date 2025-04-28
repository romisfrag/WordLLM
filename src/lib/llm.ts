import { ChatOpenAI } from "@langchain/openai";

// Interface for model data
interface ModelData {
    id: string;
    name: string;
    description?: string;
}

// Function to fetch available models from the API
export async function fetchAvailableModels(baseURL: string, apiKey: string): Promise<string[]> {
    try {
        const response = await fetch(`${baseURL}/models`, {
            headers: {
                "Authorization": `Bearer ${apiKey}`,
                "HTTP-Referer": "https://localhost:3000",
                "X-Title": "WordLLM",
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to fetch models: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();
        return data.data.map((model: ModelData) => model.id);
    } catch (error) {
        console.error("Error fetching models:", error);
        throw error;
    }
}

// Function to filter models based on search term
export function filterModels(models: string[], searchTerm: string): string[] {
    const term = searchTerm.toLowerCase();
    return models.filter(model => model.toLowerCase().includes(term));
}

// Function to initialize the model with custom configuration
export function initializeModel(baseURL: string, openAIApiKey: string, modelName?: string) {
    return new ChatOpenAI({
        openAIApiKey: openAIApiKey,
        configuration: {
            baseURL: baseURL,
            defaultHeaders: {
                "HTTP-Referer": "https://localhost:3000",
                "X-Title": "WordLLM",
            },
        },
        modelName: modelName || "qwen/qwen-2.5-7b-instruct:free",
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