/* global Office */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    // eslint-disable-next-line no-undef
    item = Office.context.mailbox.item;
  }
});