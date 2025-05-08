import { getLLMResponse } from "./llm";
import { marked } from 'marked';
import { ChatOpenAI } from "@langchain/openai";
import { initializeModel } from "./llm";
import { getFromLocalStorage } from "./local_storage";


async function concatenatePrompt(promptUrl: string, selectedText: string, isReplaceSelection: boolean = false) {
  if (promptUrl === "") {
    return selectedText;
  } else {
    const response = await fetch(promptUrl);
    if (!response.ok) {
      throw new Error(`Failed to load prompt: ${response.status} ${response.statusText}`);
    }
    const promptText = await response.text();
    // Remove trailing colon if present
    
    if (isReplaceSelection) {
      const cleanPromptText = promptText.replace(/:$/, '.');
      // Add HTML preservation prompt only for replace selection
      const htmlPreserveResponse = await fetch('/prompts/html_preserve.txt');
      if (!htmlPreserveResponse.ok) {
        throw new Error(`Failed to load HTML preservation prompt: ${htmlPreserveResponse.status} ${htmlPreserveResponse.statusText}`);
      }
      const htmlPreserveText = await htmlPreserveResponse.text();
      return cleanPromptText + ' ' + htmlPreserveText + '\n' + selectedText;
    } else {
      return promptText + '\n' + selectedText;
    }
  } 
}

export async function askLLMUrlPrompt(userText: string, taskPane: boolean = false, promptUrl: string = "", model: ChatOpenAI) {
  const fullPrompt = await concatenatePrompt(promptUrl, userText, !taskPane);
  console.log('Full prompt:', fullPrompt);
  
  // Extract prompt name from URL if available
  let promptName;
  if (promptUrl) {
    const urlParts = promptUrl.split('/');
    promptName = urlParts[urlParts.length - 1].replace('.txt', '');
  }
  
  // Get configuration from local storage
  const baseURL = getFromLocalStorage('baseURL');
  const apiKey = getFromLocalStorage('apiKey');
  const selectedModel = getFromLocalStorage('selectedModel');
  
  // Create a new model instance with the prompt name
  const modelWithPrompt = initializeModel(
    baseURL,
    apiKey,
    selectedModel,
    promptName
  );
  
  askLLM(fullPrompt, taskPane, modelWithPrompt);
}

export async function askLLMStrPrompt(userText: string, taskPane: boolean = false, promptStr: string = "", model: ChatOpenAI) {
  const fullPrompt = promptStr + "\n" + userText;
  console.log('Full prompt:', fullPrompt);
  askLLM(fullPrompt, taskPane, model);
}


// If taskPane is true, the response is displayed in the taskpane
// If taskPane is false, the response replace the selection
export async function askLLM(prompt: string, taskPane: boolean = false, model: ChatOpenAI) {
    // Show loading overlay
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = 'flex';
        overlay.classList.add('active');
    }

    try {
        await Word.run(async (context) => {
            // Get response from LLM
            const llmResponse = await getLLMResponse(prompt, model);
            console.log('Received the LLM response:', llmResponse);

            // Display the response in the taskpane
            if (taskPane) {
                const responseDiv = document.getElementById('response');
                if (responseDiv) {
                    const parsedMarkdown = await marked.parse(llmResponse);
                    console.log('Displaying response in taskpane : \n ' + parsedMarkdown);
                    responseDiv.innerHTML = `<div class="markdown-content">${parsedMarkdown}</div>`;
                    
                    // Scroll to the response section
                    const responseSection = document.querySelector('.response-section');
                    if (responseSection) {
                        responseSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                }
            } else {
                // Get the current selection's HTML
                const selection = context.document.getSelection();
                const html = selection.getHtml();
                await context.sync();
                
                // Replace the body content with the LLM response
                const bodyRegex = /<body[^>]*>([\s\S]*)<\/body>/i;
                const modifiedHtml = html.value.replace(bodyRegex, `<body>${llmResponse}</body>`);
                
                // Clear the selection first
                selection.clear();
                
                // Insert the modified HTML
                selection.insertHtml(modifiedHtml, Word.InsertLocation.replace);
            }

            await context.sync();
        });
    } catch (error) {
        console.error('Error in askLLM:', error);
        const responseDiv = document.getElementById('response');
        if (responseDiv) {
            console.log('Displaying error in UI');
            responseDiv.textContent = `An error occurred: ${error.message || error}. Please try again.`;
            
            // Scroll to the response section even on error
            const responseSection = document.querySelector('.response-section');
            if (responseSection) {
                responseSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    } finally {
        // Hide loading overlay
        if (overlay) {
            overlay.style.display = 'none';
            overlay.classList.remove('active');
        }
    }
    console.log('Finished askLLM function');
}

