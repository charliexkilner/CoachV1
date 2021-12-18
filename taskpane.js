/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

$("#run").click(() => tryCatch(run));

function run() {
  return Word.run(function(context) {
    var range = context.document.getSelection();
    range.font.color = "red";
    range.load("text");

    return context.sync().then(function() {
      console.log('The selected text was "' + range.text + '".');
    });
  });
}

/** Default helper for invoking an action and handling errors. */
function tryCatch(callback) {
  Promise.resolve()
    .then(callback)
    .catch(function(error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    });
}

const OpenAI = require('openai-api');

// Load your key from an environment variable or secret management service
// (do not include your key directly in your code)
const OPENAI_API_KEY = process.env.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;

const openai = new OpenAI(OPENAI_API_KEY);

(async () => {
  const gptResponse = await openai.complete({
      engine: 'davinci',
      prompt: 'this is a test',
      maxTokens: 5,
      temperature: 0.9,
      topP: 1,
      presencePenalty: 0,
      frequencyPenalty: 0,
      bestOf: 1,
      n: 1,
      stream: false,
      stop: ['\n', "testing"]
  });

  console.log(gptResponse.data);
})();
