import { getLLMResponse } from "./llm";

async function concatenatePrompt(promptUrl: string, selectedText: string) {
  if (promptUrl === "") {
    return selectedText;
  } else {
    const response = await fetch(promptUrl);
    if (!response.ok) {
      throw new Error(`Failed to load prompt: ${response.status} ${response.statusText}`);
    }
    const promptText = await response.text();
    return promptText + '\n' + selectedText;
  }
}

// If taskPane is true, the response is displayed in the taskpane
// If taskPane is false, the response replace the selection
export async function askLLM(userText: string, taskPane: boolean = false, promptUrl: string = "") {
    try {
        await Word.run(async (context) => {
            // Get the current selection
            
            const fullPrompt = await concatenatePrompt(promptUrl, userText);
            console.log('Full prompt:', fullPrompt);

            // Get response from LLM
            const llmResponse = await getLLMResponse(fullPrompt);
            console.log('Received LLM response:', llmResponse);

            // Display the response in the taskpane
            if (taskPane) {
              const responseDiv = document.getElementById('response');
              if (responseDiv) {
                  responseDiv.textContent = llmResponse;
              }
            } else {
              const selection = context.document.getSelection();
              selection.insertText(llmResponse, Word.InsertLocation.replace);
            }

            await context.sync();
        });
    } catch (error) {
        console.error('Error in askLLM:', error);
        const responseDiv = document.getElementById('response');
        if (responseDiv) {
            console.log('Displaying error in UI');
            responseDiv.textContent = `An error occurred: ${error.message || error}. Please try again.`;
        }
    }
    console.log('Finished askLLM function');
}

