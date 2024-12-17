"/* @OnlyCurrentDoc */"
// ENTRANCE FUNCTION
function rfpcomplete(cellReference, option, product) {
  if (!cellReference || typeof cellReference !== 'string') {
    return "Invalid or missing cell reference. Please provide a valid cell reference in A1 notation.";
  }
  var sheet = SpreadsheetApp.getActiveSheet();
  var prompt = constructPrompt(cellReference, option, product);
  var response = callOpenAI(prompt);
  return response;
}
// CONSTRUCT PROMPT
function constructPrompt(cellData, option, product) {
  var prompt;
  switch(option) {
    case "Boolean":
      // For Boolean, the response should be true or false.
      prompt = "Answer with a Boolean, true or false. Return only the answer:" + cellData;
      break;
    case "Binary":
      // For Binary, the response should be yes or no.
      prompt = "Answer with 'yes' or 'no'. Return only the answer: " + cellData;
      break;
    case "Short":
      // For a short paragraph, prompt for a concise explanation.
      prompt = "Provide a 1-2 sentence explanation for: " + cellData;
      break;
    case "Long":
      // For a long paragraph, prompt for a detailed explanation.
      prompt = "Provide a paragraph explanation for: " + cellData;
      break;
    default:
      // Default case if the option doesn't match any category.
      prompt = "Provide a 1-2 sentence explanation for: " + cellData;
  }
  if (product) {
    var expandedPrompt = "Your answer should be in the context of " + product + " " + prompt
    return expandedPrompt
  } else {
    return prompt;
  }
}
var scriptProps = PropertiesService.getScriptProperties()
// CALL OPENAI
function callOpenAI(prompt) {
  var apiKey = scriptProps.getProperty('OPENAI_API_KEY');
  var apiUrl = "https://openai-proxy.shopify.ai/v2/openai/v1/chat/completions";
  var payload = {
    "model": "openai:gpt-4o-2024-08-06",
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
