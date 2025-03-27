/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* eslint-disable prettier/prettier */
/* global Office, Word */

// State management (necessary improvement)
const state = {
  errors: [],
  currentIndex: 0
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    try {
      // Function registration with error handling (necessary)
      Office.actions.associate("checkDocumentText", checkDocumentText);
      Office.actions.associate("acceptAllChanges", acceptAllChanges);
      Office.actions.associate("rejectAllChanges", rejectAllChanges);
      Office.actions.associate("acceptCurrentChange", acceptCurrentChange);
      Office.actions.associate("rejectCurrentChange", rejectCurrentChange);
    } catch (error) {
      console.error("Function registration failed:", error);
    }
  }
});

function determineCorrectPreposition(word) {
  if (!word) return null;

  const unvoicedConsonants = new Set(['c', 'č', 'f', 'h', 'k', 'p', 's', 'š', 't']);
  let firstChar = "";
  
  for (const char of word) {
    if (char.match(/[a-zA-ZčČšŠžŽ]/)) {
      firstChar = char.toLowerCase();
      break;
    }
  }
  
  return firstChar ? (unvoicedConsonants.has(firstChar) ? "s" : "z") : null;
}

async function checkDocumentText() {
  try {
    await Word.run(async (context) => {
      // Clear previous highlights
      state.errors.forEach(err => {
        err.range.font.highlightColor = null;
      });
      state.errors = [];
      state.currentIndex = 0;

      // Search for prepositions (original logic)
      const searchOptions = { matchCase: false, matchWholeWord: true };
      const sResults = context.document.body.search("s", searchOptions);
      const zResults = context.document.body.search("z", searchOptions);
      sResults.load("items");
      zResults.load("items");
      await context.sync();

      // Process results
      const errors = [...sResults.items, ...zResults.items]
        .filter(prep => ['s', 'z'].includes(prep.text.trim().toLowerCase()))
        .map(prep => ({
          prepositionRange: prep,
          nextWordRange: prep.getNextTextRange("Word")
        }))
        .filter(candidate => {
          candidate.nextWordRange.load("text");
          return true;
        });

      await context.sync();

      // Validate and highlight errors
      state.errors = errors
        .map(({prepositionRange, nextWordRange}) => {
          const currentPrep = prepositionRange.text.trim().toLowerCase();
          const correctPrep = determineCorrectPreposition(nextWordRange.text.trim());
          return correctPrep && currentPrep !== correctPrep ? {
            range: prepositionRange,
            suggestion: correctPrep
          } : null;
        })
        .filter(Boolean);

      state.errors.forEach(err => {
        err.range.font.highlightColor = "Yellow";
      });

      await context.sync();

      if (state.errors.length > 0) {
        state.errors[0].range.select();
      } else {
        context.document.body.insertComment("No preposition errors found.", "start");
      }
    });
  } catch (error) {
    console.error("Document check failed:", error);
  }
}

// Correction functions with error handling (necessary)
async function acceptCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  
  try {
    await Word.run(async (context) => {
      const err = state.errors[state.currentIndex];
      err.range.insertText(err.suggestion, Word.InsertLocation.replace);
      err.range.font.highlightColor = null;
      await context.sync();
      state.currentIndex++;
      
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (error) {
    console.error("Failed to accept change:", error);
  }
}

async function rejectCurrentChange() {
  if (state.currentIndex >= state.errors.length) return;
  
  try {
    await Word.run(async (context) => {
      const err = state.errors[state.currentIndex];
      err.range.font.highlightColor = null;
      await context.sync();
      state.currentIndex++;
      
      if (state.currentIndex < state.errors.length) {
        state.errors[state.currentIndex].range.select();
      }
    });
  } catch (error) {
    console.error("Failed to reject change:", error);
  }
}

// Bulk operations with error handling (necessary)
async function acceptAllChanges() {
  if (state.errors.length === 0) return;
  
  try {
    await Word.run(async (context) => {
      for (const err of state.errors) {
        err.range.insertText(err.suggestion, Word.InsertLocation.replace);
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (error) {
    console.error("Failed to accept all changes:", error);
  }
}

async function rejectAllChanges() {
  if (state.errors.length === 0) return;
  
  try {
    await Word.run(async (context) => {
      for (const err of state.errors) {
        err.range.font.highlightColor = null;
      }
      await context.sync();
      state.errors = [];
    });
  } catch (error) {
    console.error("Failed to reject all changes:", error);
  }
}

// Maintain window exports for HTML buttons
window.checkDocumentText = checkDocumentText;
window.acceptAllChanges = acceptAllChanges;
window.rejectAllChanges = rejectAllChanges;
window.acceptCurrentChange = acceptCurrentChange;
window.rejectCurrentChange = rejectCurrentChange;