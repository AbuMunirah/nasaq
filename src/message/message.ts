/*
 * Dialog that shows a Nasaq message passed in the "msg" query parameter.
 */

/* global document Office URLSearchParams window */

Office.onReady(() => {
  const message = new URLSearchParams(window.location.search).get("msg") ?? "";
  const arabic = (Office.context.displayLanguage ?? "").toLowerCase().startsWith("ar");
  document.documentElement.lang = arabic ? "ar" : "en";
  document.documentElement.dir = arabic ? "rtl" : "ltr";
  document.getElementById("message")!.textContent = message;
  document.getElementById("ok")!.textContent = arabic ? "موافق" : "OK";
  document
    .getElementById("ok")!
    .addEventListener("click", () => Office.context.ui.messageParent("close"));
});
