// Run: node test.js
var assert = require("assert");
var buildSignature = require("./src/commands.js").buildSignature;

function text(html) { return html.replace(/<[^>]+>/g, "|").replace(/&nbsp;/g, " ").split("|").filter(Boolean).join("\n"); }
function sig(email, isReply, extra) {
  return buildSignature(Object.assign({ name: "Jane Doe", pronouns: "she/her", title: "Partner", extension: "1234", email: email }, extra), isReply);
}

// Text must match example-signatures/*.txt
assert.strictEqual(text(sig("jdoe@revealdesigngroup.com", false)),
  "Jane Doe\n (she/her)\nPartner\n \nreveal design group\n104 West 27\nth\n Street, 2\nnd\n Floor\nNew York, NY 10001\n(p) 212-633-4353 (x1234)\nhttp://revealdesigngroup.com\nhttps://www.instagram.com/revealdesigngrp\n ");
assert.strictEqual(text(sig("jdoe@revealdesigngroup.com", true)),
  "Jane Doe\n (she/her)\nPartner\n \nreveal design group\n(p) 212-633-4353 (x1234)\n ");
assert.strictEqual(text(sig("jdoe@CODA-LP.com", false)),
  "Jane Doe\n (she/her)\nPartner\n \nCoda lighting procurement\n104 West 27\nth\n Street, 2\nnd\n Floor\nNew York, NY 10001\n(p) 212-633-4356 (x1234)\nhttp://coda-lp.com\n ");
assert.strictEqual(text(sig("jdoe@coda-lp.com", true)),
  "Jane Doe\n (she/her)\nPartner\n \nCoda lighting procurement\n(p) 212-633-4356 (x1234)\n ");

// Optional fields drop cleanly
var bare = text(sig("jdoe@coda-lp.com", true, { pronouns: "", title: "", extension: "" }));
assert.strictEqual(bare, "Jane Doe\n \nCoda lighting procurement\n(p) 212-633-4356\n ");

assert.strictEqual(sig("jdoe@gmail.com", false), "", "other domain: none");
assert.strictEqual(sig("", false), "", "no email: none");
assert(!sig("a@coda-lp.com", false, { title: "<b>x</b>" }).includes("<b>x"), "user input escaped");

console.log("ok");
