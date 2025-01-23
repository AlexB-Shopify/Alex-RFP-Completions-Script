"/* @OnlyCurrentDoc */"

/**
 * Custom function to complete RFP parameters.
 *
 * @param {Range} cellReference - A reference to a cell in the sheet.
 * @param {string} option - A string representing the four answer types which are Boolean, Binary, Short, Long (optional)
 * @param {string} product - A string representing the specific product of interest (optional)
 * @param {string} instructions - Additional instructions or preferences to consider (optional)
 * @return {string} - The prompt that will be sent to ChatGPT.
 *
 * Example usage:
 * =rfpcomplete(A1, "Short", "Shopify Payments", "Exactitude is very important, answer NEED MORE DETAILS if unsure or if the question is poorly formulated.")
 */

// ENTRANCE FUNCTION
function rfpcomplete(cellReference, option, product, instructions) {
  if (!cellReference || typeof cellReference !== 'string') {
    return "Invalid or missing cell reference. Please provide a valid cell reference in A1 notation.";
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var prompt = constructPrompt(cellReference, option, product, instructions);
  var response = callOpenAI(prompt);
  return response;
}

// CONSTRUCT PROMPT
function constructPrompt(cellData, option, product, instructions) {
  var rfpPrompt;
  
  var contextInit = "CONTEXT: You are an eCommerce expert, and have an in-depth understanding of SHOPIFY, including its current offerings and limitations. Your addressing a merchant that is evaluating the Shopify platform, and your answers will be used to fill said merchant's Request For Proposal (RFP). "
  
  var instructionsInit = "INSTRUCTIONS: Answer in the 3rd person (e.g.:'Shopify is...'), OR use 'we' (as if you were part of Shopify). ";

  switch(option) {
    case "Boolean":
      // For Boolean, the response should be true or false.
      rfpPrompt = "Answer with a Boolean, true or false. Return only the answer: " + cellData;
      break;
    case "Binary":
      // For Binary, the response should be yes or no.
      rfpPrompt = "Answer with yes or no. Return only the answer: " + cellData;
      break;
    case "Short":
      // For a short paragraph, prompt for a concise explanation.
      rfpPrompt = "Provide a 1-2 sentence explanation for: " + cellData;
      break;
    case "Long":
      // For a long paragraph, prompt for a detailed explanation.
      rfpPrompt = "Provide a paragraph explanation for: " + cellData;
      break;
    default:
      // Default case if the option doesn't match any category.
      rfpPrompt = "Answer the question DIRECTLY and PROMPTLY. Although your answer should (reasonably) match the length and tone of the question, shorter is preferred." + "Here is the RFP point/question to address: " + cellData;
  }
  if (product) {
    var expandedPrompt = contextInit + "Your answer should be in the context of " + product + " " + instructionsInit + rfpPrompt;
    return expandedPrompt;
  } else if (instructions) {
    var expandedPrompt = contextInit + instructionsInit + "ADDITIONAL INSTRUCTIONS: " + instructions + " " + rfpPrompt;
    return expandedPrompt;
  } else if (product && instructions) {
    var expandedPrompt = contextInit + "Your answer should be in the context of " + product + " " + instructionsInit + " " + "ADDITIONAL INSTRUCTIONS: " + instructions + " " + rfpPrompt;
    return expandedPrompt;
  } else {
    rfpPrompt = contextInit + instructionsInit + rfpPrompt;
    return rfpPrompt;
  }
}

var scriptProps = PropertiesService.getScriptProperties()
// CALL OPENAI

function callOpenAI(prompt) {
  var apiKey = scriptProps.getProperty('OPENAI_API_KEY');
  var apiUrl = "[INSERT API URL HERE]";
  var payload = {
    "model": "gpt-4o-2024-08-06",
    "messages": [{
      "role": "user",
      "content": prompt
    }],
    "max_tokens": 300
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  var response = UrlFetchApp.fetch(apiUrl, options);
  var content = JSON.parse(response.getContentText());
  if (response.getResponseCode() === 200) {
    var messages = content.choices;
    var lastBotMessage = messages[messages.length - 1].message;
    return lastBotMessage.content.trim();
  } else {
    if (content.error) {
      throw new Error(content.error.message);
    }
    throw new Error("Failed to fetch response from OpenAI.");
  }
}
