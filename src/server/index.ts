import {
  onOpen,
  openSendEmailDialog,
  testMailMerge,
  testMassSend
} from "./ui";

import { sendEmails } from "./mailMerge";
import { sendMassEmail } from "./massSend";

// Public functions must be exported as named exports

// Expose UI functions
export {
  sendEmails,
  sendMassEmail,
  openSendEmailDialog,
  testMailMerge,
  testMassSend,
  onOpen,
};
