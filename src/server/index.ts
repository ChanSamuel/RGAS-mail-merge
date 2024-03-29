import {
  onOpen,
  openSendEmailDialog,
  testMailMerge,
  testMassSend
} from "./ui";

import { getSheetsData, addSheet, deleteSheet, setActiveSheet, getColumnValuesByName, getColumnNames } from "./helpers/sheets";
import { createEmailSentColsIfNotExists } from "./helpers/general"
import { createDraftWithGmailAPI, getGmailTemplateFromDrafts } from "./helpers/gmail";
import { sendEmails } from "./mailMerge";
import { sendMassEmail } from "./massSend";

// Public functions must be exported as named exports

// Expose UI functions
export {
  onOpen,
  openSendEmailDialog,
  testMailMerge,
  testMassSend,
};

// Expose Sheets helper functions
export {
  getSheetsData,
  addSheet,
  deleteSheet,
  setActiveSheet,
  getColumnValuesByName,
  getColumnNames
}

// Expose General helper functions
export {
  createEmailSentColsIfNotExists
};

// Expose Gmail helper functions
export {
  createDraftWithGmailAPI,
  getGmailTemplateFromDrafts
};

// Expose Mail Merge functions
export {
  sendEmails
};

// Expose Mass Send functions
export {
  sendMassEmail
};
