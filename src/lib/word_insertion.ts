import { getLLMResponse } from "./llm";

export async function sendPrompt() {
    console.log('Starting sendPrompt function');
    try {
      await Word.run(async (context) => {
        // Get the prompt from the textarea
        const prompt = (document.getElementById('prompt') as HTMLTextAreaElement).value;
        console.log('Retrieved prompt:', prompt);
              
        // Get response from LLM
        const response = await getLLMResponse(prompt);
        console.log('Received LLM response:', response);
        
        // Insert the response into the document
        const paragraph = context.document.body.insertParagraph(response, Word.InsertLocation.end);
  
        await context.sync();
      });
    } catch (error) {
      console.error('Error in sendPrompt:', error);
      const responseDiv = document.getElementById('response');
      if (responseDiv) {
        console.log('Displaying error in UI');
        responseDiv.textContent = `An error occurred: ${error.message || error}. Please try again.`;
      }
    }
    console.log('Finished sendPrompt function');
  }