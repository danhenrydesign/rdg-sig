// Event handlers for OnNewMessageCompose and OnMessageFromChanged.
// ES5 on purpose: classic Outlook on Windows runs this file in a JavaScript-only runtime (no DOM, no modules).

// Per-domain text, copied from example-signatures/. The two brands share a layout and differ only in this text.
var BRANDS = {
  "revealdesigngroup.com": {
    company: "reveal design group",
    phone: "212-633-4353",
    links: ["http://revealdesigngroup.com", "https://www.instagram.com/revealdesigngrp"]
  },
  "coda-lp.com": {
    company: "Coda lighting procurement",
    phone: "212-633-4356",
    links: ["http://coda-lp.com"]
  }
};
var ADDRESS = ["104 West 27<sup>th</sup> Street, 2<sup>nd</sup> Floor", "New York, NY 10001"];

var FONT_LT = 'font-family:&quot;Angostura Lt&quot;,sans-serif;';
var FONT_BL = 'font-family:&quot;Angostura Bl&quot;,sans-serif;';
var GREY = "color:#404040;";
var ORANGE = "color:#E36C0A;";

function domainOf(email) {
  return String(email || "").toLowerCase().split("@")[1] || "";
}

function esc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function p(html) {
  return '<p style="margin:0;font-size:12pt;' + FONT_LT + GREY + '">' + html + "</p>";
}

// user: { name, email, pronouns, title, extension }. Returns "" when the domain has no signature.
function buildSignature(user, isReply) {
  var brand = BRANDS[domainOf(user.email)];
  if (!brand) return "";
  var lines = [];

  lines.push(p("<b>" + esc(user.name) + "</b>" + (user.pronouns ? " (" + esc(user.pronouns) + ")" : "")));
  if (user.title) lines.push(p(esc(user.title)));
  lines.push(p("&nbsp;"));
  lines.push(p('<b><span style="' + FONT_BL + ORANGE + '">' + esc(brand.company) + "</span></b>"));
  if (!isReply) lines = lines.concat(ADDRESS.map(p));
  lines.push(p("(p) " + brand.phone + (user.extension ? " (x" + esc(user.extension) + ")" : "")));
  if (!isReply) {
    lines = lines.concat(brand.links.map(function (url) {
      return p('<a href="' + url + '" style="color:blue">' + url + "</a>");
    }));
  }
  lines.push(p("&nbsp;"));
  return "<div>" + lines.join("") + "</div>";
}

function applySignature(event) {
  var item = Office.context.mailbox.item;
  var done = function () { event.completed(); };

  item.from.getAsync(function (fromResult) {
    // From is the alias the user picked, which is what the signature should show.
    var email = fromResult.status === Office.AsyncResultStatus.Succeeded && fromResult.value ? fromResult.value.emailAddress : "";
    item.getComposeTypeAsync(function (typeResult) {
      var isReply = typeResult.status === Office.AsyncResultStatus.Succeeded && typeResult.value.composeType !== "newMail";
      var saved = Office.context.roamingSettings.get("profile") || {};
      var html = buildSignature({
        name: Office.context.mailbox.userProfile.displayName,
        email: email,
        pronouns: saved.pronouns,
        title: saved.title,
        extension: saved.extension
      }, isReply);

      if (!html) {
        // Other domain: clear any signature we inserted for a previous From address.
        item.body.setSignatureAsync("", { coercionType: Office.CoercionType.Html }, done);
        return;
      }
      item.disableClientSignatureAsync(function () {
        item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, done);
      });
    });
  });
}

if (typeof Office !== "undefined" && Office.actions) {
  Office.actions.associate("onNewMessageComposeHandler", applySignature);
  Office.actions.associate("onMessageFromChangedHandler", applySignature);
}

if (typeof module !== "undefined") module.exports = { buildSignature: buildSignature };
