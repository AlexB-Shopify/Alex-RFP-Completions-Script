# Alex Beeson's RFP Completions

This Google Apps Script allows us to utilize OpenAI's capabilities to fill out customer questionnaires and RFPs directly within Google Sheets. 

It is a fork of some of the specific prompt code from a startup I used to have called Lightning Response.

By using this script, you can generate responses based on various input types (Boolean, Binary, Short, Long) for your specific needs.

## Setup Instructions

### Step 1: Open Google Sheets

1. Open a Google Sheets document where you want to use this script.

### Step 2: Access Apps Script

1. Click on `Extensions` in the menu bar.
2. Select `Apps Script`.

### Step 3: Add the Script

1. Delete any code in the script editor and copy-paste the provided Google Apps Script code into the editor.
2. Save the project with a name of your choice.

### Step 4: Get An API Key
1. Head over to our Open AI Proxy site: https://openai-proxy.shopify.io/
2. Click on "Personal Access" from the menu.
3. Ensure OpenAI is selected.
4. Copy the key in the box above.
5. Reminder - Key's expire every 24 hours, simply fetch a new one each day.

### Step 5: Set Up API Key

1. Navigate to the `Properties` menu in the Apps Script editor.
2. Click on `Script properties`.
3. Add a new property:
   - **Key**: `OPENAI_API_KEY`
   - **Value**: Your OpenAI Proxy API key (obtained above).

### Step 6: Authorize the Script

1. Click on the `Run` button (▶️) in the Apps Script editor.
2. You will be prompted to authorize the script. Follow the on-screen instructions to give the necessary permissions.

## Using the Script in Google Sheets

To use the `rfpcomplete` function in your Google Sheets, follow these steps:

1. **Select a Cell**: Choose the cell where you want the output to appear.
2. **Enter the Function**: Use the following format to call the function:

   ```
   =rfpcomplete("A1", "Option", "Product")
   ```

   - **cellReference**: The cell (in A1 notation) containing the question or prompt.
   - **option**: The type of response you want:
     - `"Boolean"` for true/false answers
     - `"Binary"` for yes/no answers
     - `"Short"` for a brief explanation
     - `"Long"` for a detailed explanation
   - **product**: (Optional) The context of the product related to the question.

### Example Usage

If you have a question in cell A1 and want a short response regarding a specific product:

```
=rfpcomplete("A1", "Short", "Shopify Payments")
```

## Roadmap
I am currently working with some of our Vaultbot engineers to improve it's API and the customization of experts, we will then add additional parameters to our script and point requests at Vaultbot for embedded answers. This will focus the context window on up-to-date answers and fewer hallucinations.

## Troubleshooting

If you encounter any issues:

- Check that your API key is correctly set in the script properties.
- Check that your API key hasn't expired
- Ensure that the prompt and options are correctly formatted in your function call.
- Hover over the errored cell for additional detail on the error message