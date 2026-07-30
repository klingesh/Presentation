"""UNIT 4 — Ethics, legal and social implications."""

from deck import (Deck, notes, BLUE, TEAL, AMBER, VIOLET, GREEN, RED)

OUT = "Unit 4 - Ethics, Legal and Social Implications.pptx"


def build(outdir="."):
    d = Deck(4, "Ethics, Legal and Social Implications")

    # ---------------------------------------------------------------- 01
    s = d.title_slide(
        "Ethics, Privacy and Regulation",
        "The rules, risks and responsibilities that come with using AI - and how "
        "to keep a human properly in charge of decisions.",

        chips=["Privacy", "Security", "Regulation", "Decision Making"])
    notes(s, "Frame this positively: these are not obstacles. They are what make "
             "AI safe enough to actually deploy at scale.")

    # ---------------------------------------------------------------- 02
    d.agenda_slide([
        ("Why ethics is practical", "It is a business risk, not a philosophy class"),
        ("Privacy", "What data you may collect, keep and use"),
        ("Security", "New attack surfaces that come with AI"),
        ("Bias and fairness", "How unfairness gets in, and how to catch it"),
        ("AI regulations", "The global picture in plain terms"),
        ("Risk-based rules", "Why the same tool is fine here and banned there"),
        ("AI in decision making", "Who decides, and who is answerable"),
        ("A governance checklist", "What good practice looks like day to day"),
    ], note="One rule ties the whole unit together: a person is always "
            "accountable for the outcome.")

    # ---------------------------------------------------------------- 03
    d.section_slide(1, "Why This Matters Commercially",
                    "Ethics failures show up as fines, lawsuits and lost "
                    "customers.", BLUE)

    # ---------------------------------------------------------------- 04
    s = d.cards_slide(
        "The business case", "Four reasons this is not optional",
        [("Legal exposure", "Fines under data and AI law are now a percentage of "
                            "global revenue, not a small penalty."),
         ("Reputation", "One unfair or leaked outcome can undo years of brand "
                        "building."),
         ("Customer trust", "People share data with organisations they believe "
                            "will handle it properly."),
         ("Access to markets", "Buyers and regulators increasingly ask how your "
                               "AI works before they sign.")],
        sub="Doing this well is a commercial advantage, not a cost centre.",
        cols=4, accent=BLUE, tinted=True)

    # ---------------------------------------------------------------- 05
    d.section_slide(2, "Privacy and Security",
                    "Protecting the data, and protecting the system.", VIOLET)

    # ---------------------------------------------------------------- 06
    s = d.cards_slide(
        "Privacy", "Six principles that cover most of the law",
        [("Purpose", "Collect data for a stated reason and use it only for that."),
         ("Minimisation", "Collect the least you need. Extra data is extra risk."),
         ("Consent", "People should know and agree, in language they understand."),
         ("Accuracy", "Wrong data produces wrong decisions about real people."),
         ("Retention", "Delete it when the purpose ends. Do not keep it forever."),
         ("Rights", "People can ask what you hold, correct it and have it "
                    "removed.")],
        sub="These appear in GDPR, India's DPDP Act and most other privacy laws.",
        cols=3, accent=VIOLET)

    # ---------------------------------------------------------------- 07
    s = d.bullets_slide(
        "The everyday risk", "How private data actually leaks with Gen AI",
        [("Pasting into public tools", "Staff paste customer data or contracts "
                                       "into a free chatbot to save time."),
         ("Training on your input", "Some consumer tools may use what you type to "
                                    "improve the model."),
         ("Output that remembers", "A model can repeat sensitive detail it saw "
                                   "during training."),
         ("Screenshots and exports", "Data leaves your controlled systems and "
                                     "lands in someone's downloads folder."),
         ("Shadow AI", "Teams adopt tools nobody approved, so nobody can secure "
                       "them."),
         ("Third-party add-ons", "A browser plugin can read everything on the "
                                 "page, including customer records.")],
        sub="Most incidents are not attacks. They are ordinary people taking a "
            "shortcut.",
        lead="The single most useful policy: never paste anything into a public AI "
             "tool that you would not post publicly.",
        accent=RED, two_col=True)

    # ---------------------------------------------------------------- 08
    s = d.cards_slide(
        "Security", "New risks that come with AI systems",
        [("Prompt injection", "Hidden instructions inside a document or web page "
                              "hijack what the AI does next."),
         ("Data poisoning", "Bad data is fed in deliberately so the model learns "
                            "the wrong thing."),
         ("Model theft", "Repeated querying is used to copy the behaviour of a "
                         "model you paid to build."),
         ("Deepfakes", "Cloned voice or video used to authorise a payment or "
                       "impersonate a leader."),
         ("Over-permissioned agents", "An agent connected to too many systems can "
                                      "do too much damage."),
         ("Supply chain", "A vulnerability in a third-party AI service becomes "
                          "your incident.")],
        sub="Old security still applies. These are the additions.",
        cols=3, accent=RED)
    notes(s, "Deepfake voice fraud is the one to highlight - finance teams have "
             "authorised large transfers on a call that sounded exactly like the "
             "CFO. Verification process beats voice recognition.")

    # ---------------------------------------------------------------- 09
    s = d.process_slide(
        "Practical protection", "Five controls that prevent most problems",
        [("Approve tools", "Publish a short list of tools cleared for company data."),
         ("Classify data", "Mark what may never leave your own systems."),
         ("Train people", "One hour of clear guidance prevents most incidents."),
         ("Limit access", "Agents and staff see only what their role needs."),
         ("Log and review", "Keep records so you can investigate and improve.")],
        accent=[BLUE, TEAL, AMBER, VIOLET, GREEN],
        note="Banning AI does not work - people use it on their phones instead. "
             "Provide a safe approved option.")

    # ---------------------------------------------------------------- 10
    d.section_slide(3, "Bias and Fairness",
                    "The risk that is hardest to see and easiest to inherit.",
                    AMBER)

    # ---------------------------------------------------------------- 11
    s = d.split_slide(
        "How it happens", "Bias is learned, not programmed",
        "Nobody writes an unfair rule. The model copies the pattern it was shown.",
        ["Past decisions were made by people, with their biases",
         "The model learns those patterns as 'normal'",
         "Groups missing from the data get served worst",
         "The output looks neutral and statistical",
         "At scale, a small bias affects thousands of people"],
        [("Representation bias", "Some groups barely appear in the training data."),
         ("Historical bias", "The past was unfair, so the prediction is too."),
         ("Proxy bias", "A postcode or school name stands in for something "
                        "protected.")],
        accent=AMBER)

    # ---------------------------------------------------------------- 12
    s = d.quadrant_slide(
        "Managing it", "Four practical fairness controls",
        [("Test by group", ["Measure accuracy for each group separately",
                            "An overall score can hide a serious gap",
                            "Do this before launch and again later"], BLUE),
         ("Mixed review teams", ["People who differ spot different problems",
                                 "Involve the affected function, not only IT",
                                 "Document what you checked"], TEAL),
         ("Explainability", ["Be able to say which factors mattered",
                             "If you cannot explain it, do not automate it",
                             "Required by law in many high-risk uses"], VIOLET),
         ("Appeal route", ["A person can ask for a human review",
                           "Someone has authority to overturn",
                           "Track overturns - they show where it fails"], GREEN)],
        sub="You cannot remove bias completely. You can measure it and manage it.",
        accent=AMBER)

    # ---------------------------------------------------------------- 13
    d.section_slide(4, "AI Regulations",
                    "What the law now expects, in plain language.", GREEN)

    # ---------------------------------------------------------------- 14
    s = d.table_slide(
        "The global picture", "The frameworks that matter most",
        ["Framework", "Where", "What it does", "Why you care"],
        [["EU AI Act", "European Union", "Sorts AI by risk level and bans the "
                                        "worst uses", "The global benchmark - "
                                                      "others copy it"],
         ["GDPR", "European Union", "Governs personal data and automated "
                                    "decisions", "Applies if you touch EU data"],
         ["DPDP Act", "India", "Consent, purpose limits and duties for data "
                               "handlers", "Direct compliance duty in India"],
         ["NIST AI RMF", "United States", "A voluntary framework for managing AI "
                                          "risk", "Practical template for "
                                                  "governance"],
         ["Sector rules", "Everywhere", "Finance, health and hiring rules already "
                                        "apply", "Existing law did not pause for "
                                                 "AI"]],
        widths=[1.4, 1.4, 2.5, 2.3], accent=GREEN,
        note="Details differ by country. The direction is the same everywhere: "
             "more transparency and more accountability.")

    # ---------------------------------------------------------------- 15
    s = d.steps_slide(
        "The core idea", "Regulation is risk-based, not technology-based",
        [("Minimal risk", "Spam filters, grammar help, product recommendations. "
                          "Use freely.", GREEN),
         ("Limited risk", "Chatbots and generated content. Duty: tell people it "
                          "is AI.", TEAL),
         ("High risk", "Hiring, credit, education, healthcare. Duty: testing, "
                       "documentation, human oversight.", AMBER),
         ("Unacceptable", "Social scoring and manipulative systems. Banned "
                          "outright.", RED)],
        sub="The same technology can sit in any band. What matters is what you "
            "use it for.",
        accent=GREEN)
    notes(s, "Key insight for the class: the law does not regulate the model, it "
             "regulates the application. Recommending a film and deciding a loan "
             "are treated completely differently.")

    # ---------------------------------------------------------------- 16
    s = d.cards_slide(
        "Compliance in practice", "Six duties that keep coming up",
        [("Disclose AI use", "Tell people when they are interacting with AI or "
                             "seeing generated content."),
         ("Keep documentation", "Record what the system does, on what data, and "
                                "how it was tested."),
         ("Provide human oversight", "A person can review, override and stop the "
                                     "system."),
         ("Explain decisions", "Be able to give a meaningful reason for an "
                               "outcome that affects someone."),
         ("Protect the data", "Lawful basis, security controls and defined "
                              "retention."),
         ("Monitor after launch", "Performance drifts. Review it on a schedule, "
                                  "not once.")],
        cols=3, accent=BLUE)

    # ---------------------------------------------------------------- 17
    d.section_slide(5, "AI in Decision Making",
                    "The question that decides everything else: who is "
                    "accountable?", TEAL)

    # ---------------------------------------------------------------- 18
    s = d.process_slide(
        "Levels of automation", "Four ways AI can take part in a decision",
        [("Human only", "AI is not involved. The default for rare, high-stakes "
                        "calls."),
         ("AI advises", "AI recommends, a person decides. Safest useful setting."),
         ("AI decides, human reviews", "AI acts, a person checks and can "
                                       "overturn."),
         ("Fully automatic", "No human in the loop. Only for low-risk, "
                             "reversible decisions.")],
        sub="Match the level to the consequence of being wrong.",
        accent=[BLUE, TEAL, AMBER, RED],
        note="Most business decisions belong in level 2. Most incidents happen "
             "when something quietly drifts to level 4.")

    # ---------------------------------------------------------------- 19
    s = d.compare_slide(
        "Choosing the level", "Where automation is fine, and where it is not",
        {"label": "Safe to automate", "headline": "Low stakes and easy to reverse",
         "points": ["Sorting and tagging incoming requests",
                    "Suggesting products or content",
                    "Drafting internal documents",
                    "Flagging items for human review",
                    "Routine, high-volume, low-value calls"]},
        {"label": "Keep a human", "headline": "Life, money, rights or reputation",
         "points": ["Hiring, promotion and dismissal",
                    "Credit, insurance and pricing decisions",
                    "Medical and safety judgements",
                    "Anything that is hard to undo",
                    "Anything a person could reasonably appeal"]},
        lacc=GREEN, racc=RED,
        verdict="Ask one question: if this is wrong, who is harmed and can we undo "
                "it?")

    # ---------------------------------------------------------------- 20
    s = d.bullets_slide(
        "Human oversight", "What 'a human in the loop' has to mean",
        [("Real authority", "The reviewer can actually say no, without needing "
                            "permission."),
         ("Enough time", "Sixty decisions an hour is not review, it is "
                         "rubber-stamping."),
         ("Enough information", "The reviewer sees why the AI suggested this."),
         ("Named accountability", "One person owns the outcome. Not 'the "
                                  "algorithm'."),
         ("Overturns are tracked", "Frequent overrides are a signal the model "
                                   "needs work."),
         ("Training on limits", "Reviewers know where the AI is usually wrong.")],
        sub="A rubber stamp is not oversight. It is just paperwork.",
        accent=TEAL, two_col=True)

    # ---------------------------------------------------------------- 21
    s = d.quote_slide(
        "A model can make a recommendation. Only a person can be accountable.",
        "This is the line that regulators, courts and customers all care about.",
        kicker="The principle to remember")

    # ---------------------------------------------------------------- 22
    s = d.cards_slide(
        "Governance", "A practical starting checklist",
        [("Write a short AI policy", "Two pages people will actually read: "
                                     "approved tools, banned data, who to ask."),
         ("Keep an AI register", "A list of every AI use, its owner and its risk "
                                 "level."),
         ("Classify each use case", "Minimal, limited or high risk - and apply the "
                                    "matching controls."),
         ("Review on a schedule", "Quarterly checks on accuracy, fairness and "
                                  "complaints."),
         ("Train everyone once a year", "Short, practical, example-led. Not a "
                                        "legal lecture."),
         ("Have an incident route", "People know exactly who to tell when "
                                    "something goes wrong.")],
        sub="Governance is boring, cheap and the reason projects survive.",
        cols=3, accent=BLUE)

    # ---------------------------------------------------------------- 23
    s = d.takeaways_slide([
        ("Ethics is commercial risk management",
         "Fines, reputation, trust and market access."),
        ("Privacy comes down to purpose, minimisation and consent",
         "Collect less, use it only as promised, delete it on time."),
        ("Most leaks are shortcuts, not attacks",
         "Approve safe tools so people do not go around you."),
        ("Bias is inherited from data and must be measured",
         "Test by group, review with mixed teams, allow appeals."),
        ("Regulation follows risk, not technology",
         "The same model is unrestricted in one use and banned in another."),
        ("Accountability cannot be automated",
         "AI recommends. A named person decides and answers for it."),
    ], sub="Unit 4 in six lines.")

    # ---------------------------------------------------------------- 24
    d.closing_slide(
        "Thank you",
        "Next up — Unit 5: strategy, building AI-based business solutions, and "
        "how leaders make teams dramatically more effective.",
        ["What is the riskiest way AI is being used in your organisation today?",
         "Which of your decisions should never be fully automated?",
         "Who owns AI governance where you work - and do they know it?"])

    path = f"{outdir}/{OUT}"
    d.save(path)
    return path


if __name__ == "__main__":
    import sys
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
