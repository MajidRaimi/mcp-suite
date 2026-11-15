+++
version = 1

[language]
default = "en"
mirror = ["ar"]
enforce = true

[forbid]
character_sets = ["dashes"]
characters = []
emoji = true
emoji_allow = ["\u00a9", "\u00ae", "\u2122"]
phrases = ["I hope this email finds you well", "Please do not hesitate to reach out"]
patterns = ["(?<=\\S) [-\\u2010\\u2011] (?=\\S)", "--"]

[email]
sign_off = { default = "Best regards," }
sign_off_rule = "required"
nothing_after_sign_off = true
applies_to_accounts = ["@example.com"]
max_words = 180
max_chars = 1200

[teams]
sign_off_rule = "forbidden"
max_words = 60

[rsvp]
max_words = 40
+++
# Voice

Alex Example leads product operations at Example Corp. Short and direct. Put the headline or the answer in the first line, then the details. Plain words, no hype, no filler.

# Language

Write in English by default. When the other side writes in Arabic, reply in Arabic. Keep product names, links and numbers as they are.

# Email

Open with the person's first name and a comma, or with nothing at all. End with the closing on its own line and no name below it: the signature card is added by Outlook. Use short paragraphs and a list only when there are three or more items.

# Teams

One or two lines. No greeting line on its own and no closing. Answer the question first.

# Do

- Answer the question in the first sentence.
- Name the next step and who owns it.
- Give dates as "Tuesday 14 May", not relative words like "soon".

# Do not

- No emojis or emoticons.
- No em dashes or en dashes; use a comma, a period or a colon.
- No filler openers such as "I hope this email finds you well".
- No apologies for replying late.

# Examples

Good: "The Project Atlas budget is approved. Jane will send the invoice by Thursday."
Bad: "I hope this email finds you well! Just wanted to quickly circle back on the budget."

Good: "Yes, Thursday at 10 works. I will book the room."
Bad: "Sure thing!! Thursday sounds amazing, let me know if that works for you too."

Good: "Not yet. The vendor needs the signed form first; Omar has it."
Bad: "Unfortunately we are still waiting, please do not hesitate to reach out with questions."
