"""UNIT 1 — Introduction of AI and Gen AI."""

from deck import (Deck, notes, BLUE, TEAL, AMBER, VIOLET, GREEN, RED)

OUT = "Unit 1 - Introduction to AI and Gen AI.pptx"


def build(outdir="."):
    d = Deck(1, "Introduction of AI and Gen AI")

    # ---------------------------------------------------------------- 01
    s = d.title_slide(
        "Introduction to AI and Generative AI",
        "What artificial intelligence really is, the different types, how it "
        "learns, and where generative AI fits in.",

        chips=["What is AI", "Types & Models", "Gen AI vs Predictive AI",
               "Risks & Benefits"])
    notes(s, "Open with a simple hook: everyone in this room used AI today - "
             "maps, spam filters, autocomplete, ChatGPT. This unit gives us the "
             "vocabulary for the rest of the course.")

    # ---------------------------------------------------------------- 02
    s = d.agenda_slide([
        ("What is AI?", "A plain-English definition and everyday examples"),
        ("Types of AI", "By capability and by the way it works"),
        ("How AI models learn", "The four learning styles, minus the maths"),
        ("What is Generative AI?", "Machines that create, not just decide"),
        ("Gen AI vs Predictive AI", "The single most useful comparison"),
        ("Advantages", "Why every business is investing in it"),
        ("Risks and limitations", "Where it goes wrong, and why"),
        ("Using it well", "A simple rule for safe, sensible use"),
    ], note="Eight short stops. By the end you should be able to explain AI to "
            "someone with no technical background.")

    # ---------------------------------------------------------------- 03
    d.section_slide(1, "What Is Artificial Intelligence?",
                    "Starting with the basics - in normal language.", BLUE)

    # ---------------------------------------------------------------- 04
    s = d.split_slide(
        "Definition", "Artificial Intelligence",
        "Software that learns from data to do jobs that used to need a human "
        "brain.",
        ["It looks at large amounts of past data",
         "It finds patterns in that data",
         "It uses those patterns on something new",
         "It gets better as it sees more examples",
         "No one writes step-by-step rules for it"],
        [("Traditional software", "A person writes the rules. Same input always "
                                 "gives the same output."),
         ("Artificial intelligence", "The machine works out the rules from "
                                     "examples it has been shown."),
         ("Why it matters", "You can now automate judgement, not just "
                            "repetitive clicking.")],
        accent=BLUE)
    notes(s, "Key line: normal software follows rules we write. AI writes its "
             "own rules from data. That one difference explains everything else "
             "in this unit.")

    # ---------------------------------------------------------------- 05
    s = d.cards_slide(
        "Everyday examples", "You already use AI several times a day",
        [("Maps and navigation", "Predicts traffic and picks the fastest route "
                                 "using live and historic data."),
         ("Streaming and shopping", "Suggests the next show or product based on "
                                    "what people like you chose."),
         ("Bank fraud alerts", "Spots a payment that does not match your normal "
                               "spending and blocks it."),
         ("Email spam filter", "Learns from millions of flagged emails what junk "
                               "looks like."),
         ("Face unlock", "Matches the camera image to the stored pattern of "
                         "your face."),
         ("Chat assistants", "Understands your question in plain language and "
                             "writes an answer.")],
        sub="AI is not new or futuristic. It is already normal.",
        cols=3, accent=BLUE)

    # ---------------------------------------------------------------- 06
    d.section_slide(2, "Types of AI", "Two useful ways to sort it.", TEAL)

    # ---------------------------------------------------------------- 07
    s = d.cards_slide(
        "Type 1 — by capability", "How smart is it?",
        [("Narrow AI", "Very good at one job only. A chess engine cannot drive "
                       "a car.\n\nStatus: this is everything we use today."),
         ("General AI", "Would handle any task a human can, and switch between "
                        "them.\n\nStatus: does not exist yet."),
         ("Super AI", "Would beat humans at everything, including creativity.\n\n"
                      "Status: theory and debate only.")],
        sub="Almost every product marketed as 'AI' today is narrow AI.",
        cols=3, accent=TEAL, tinted=True, numbered=True)
    notes(s, "Useful myth-buster. Films show general and super AI. Business "
             "reality is narrow AI - a tool that is excellent at one narrow job.")

    # ---------------------------------------------------------------- 08
    s = d.cards_slide(
        "Type 2 — by how it works", "Four families you will hear named",
        [("Machine Learning", "Learns patterns from structured data. Used for "
                              "scores, forecasts and risk ratings."),
         ("Deep Learning", "Uses layered neural networks. Handles images, audio "
                           "and language well."),
         ("Natural Language Processing", "Reads and writes human language. Powers "
                                         "chatbots and translation."),
         ("Computer Vision", "Understands pictures and video. Used in quality "
                             "checks and security.")],
        sub="They overlap - deep learning is a branch of machine learning.",
        cols=4, accent=TEAL)

    # ---------------------------------------------------------------- 09
    s = d.process_slide(
        "How models learn", "Four ways to train a model",
        [("Supervised", "Learn from labelled examples. 'This email is spam.'"),
         ("Unsupervised", "No labels. Find natural groups in the data."),
         ("Reinforcement", "Learn by trial, reward and penalty."),
         ("Self-supervised", "Hide part of the data and predict it. Powers Gen AI.")],
        sub="A 'model' is simply the trained result - the file that holds "
            "everything the system learned.",
        accent=[BLUE, TEAL, AMBER, VIOLET],
        note="Self-supervised learning is the reason Gen AI arrived so quickly - "
             "it can learn from raw internet text without humans labelling it.")

    # ---------------------------------------------------------------- 10
    d.section_slide(3, "What Is Generative AI?",
                    "From machines that decide to machines that create.", VIOLET)

    # ---------------------------------------------------------------- 11
    s = d.split_slide(
        "Definition", "Generative AI",
        "AI that creates new content - text, images, audio, video or code - from "
        "a simple instruction.",
        ["You give it an instruction, called a prompt",
         "It predicts the most likely next word, pixel or note",
         "It repeats that prediction until the output is complete",
         "The result is new, not copied from a database",
         "Ask twice and you get two different answers"],
        [("Text", "Emails, summaries, reports, code - ChatGPT, Claude, Gemini."),
         ("Images and video", "Posters, mockups, ads - Midjourney, VEO."),
         ("Audio", "Voiceovers, music, dubbing in many languages.")],
        accent=VIOLET)
    notes(s, "Analogy that lands well: it is very advanced autocomplete. Your "
             "phone suggests the next word; Gen AI suggests the next thousand "
             "words, and keeps them consistent.")

    # ---------------------------------------------------------------- 12
    s = d.cards_slide(
        "Model families", "The main kinds of Gen AI models",
        [("Large Language Models", "Trained on huge amounts of text. Understand "
                                   "and write language. GPT, Claude, Gemini, Llama."),
         ("Diffusion models", "Start from visual noise and clean it up step by "
                              "step into an image. Midjourney, Stable Diffusion."),
         ("Multimodal models", "Accept and produce more than one format - read a "
                               "chart, answer in words."),
         ("Foundation models", "Large general models trained once, then adapted "
                               "cheaply to many specific tasks.")],
        sub="You do not need to build these. You need to know which one to pick.",
        cols=4, accent=VIOLET)

    # ---------------------------------------------------------------- 13
    d.section_slide(4, "Gen AI vs Predictive AI",
                    "The comparison that clears up most confusion.", AMBER)

    # ---------------------------------------------------------------- 14
    s = d.compare_slide(
        "Core difference", "Two different jobs, not two competitors",
        {"label": "Predictive AI", "headline": "Answers: what is likely to happen?",
         "points": ["Output is a number, score or category",
                    "Trained on your own historic data",
                    "Answer is stable and repeatable",
                    "Accuracy can be measured exactly",
                    "Example: which customers will leave next month"]},
        {"label": "Generative AI", "headline": "Answers: can you create this for me?",
         "points": ["Output is new content",
                    "Trained on very large public datasets",
                    "Answer varies each time you ask",
                    "Quality is judged by a human",
                    "Example: write the offer email to keep them"]},
        sub="They work best together, not instead of each other.",
        lacc=BLUE, racc=VIOLET,
        verdict="Predictive AI tells you what to expect. Generative AI helps you "
                "act on it.")

    # ---------------------------------------------------------------- 15
    s = d.table_slide(
        "Side by side", "Quick reference table",
        ["", "Predictive AI", "Generative AI"],
        [["Main question", "What will happen?", "What can be made?"],
         ["Output", "Score, number, label", "Text, image, audio, code"],
         ["Data used", "Company's own history", "Very large public datasets"],
         ["Same input twice", "Same answer", "Different answer"],
         ["Judged by", "Accuracy against reality", "Human review of quality"],
         ["Business use", "Forecast demand, score credit", "Draft content, "
                          "summarise, assist"],
         ["Main weakness", "Fails when the future changes", "Can state wrong "
                           "things confidently"]],
        widths=[1.5, 2.4, 2.4], accent=AMBER,
        note="If you remember one row, remember the first one.")

    # ---------------------------------------------------------------- 16
    d.section_slide(5, "Advantages, Risks and Limitations",
                    "An honest look at both sides.", GREEN)

    # ---------------------------------------------------------------- 17
    s = d.cards_slide(
        "The upside", "Why organisations are investing",
        [("Speed", "First drafts, summaries and reports in minutes instead of days."),
         ("Lower cost", "Routine work is handled by software, so teams handle "
                        "more with the same headcount."),
         ("Always available", "No shifts, no holidays, no drop in quality at "
                              "11 pm."),
         ("Consistency", "The same standard of output every single time."),
         ("Scale", "Serve ten customers or ten thousand with the same system."),
         ("Better decisions", "Patterns in large data that a person would "
                              "simply never spot.")],
        cols=3, accent=GREEN, tinted=False)

    # ---------------------------------------------------------------- 18
    s = d.quadrant_slide(
        "The downside", "Risks you must plan for",
        [("Hallucination", ["Invents facts, names and figures",
                            "Sounds completely confident while wrong",
                            "Always verify before you use it"], RED),
         ("Bias", ["Learns the bias in its training data",
                   "Can quietly treat groups unfairly",
                   "A real legal and reputation risk"], AMBER),
         ("Data leakage", ["Staff paste confidential data into public tools",
                           "That data may be stored or reused",
                           "Needs a clear company policy"], VIOLET),
         ("Over-reliance", ["Skills fade when nobody checks the output",
                            "Copy-paste answers with no thought",
                            "Keep a human accountable"], BLUE)],
        sub="None of these are reasons to avoid AI. They are reasons to use it "
            "with a process.",
        accent=RED)
    notes(s, "Hallucination is the one to stress. The model is built to produce "
             "a fluent answer, not a true one. Fluency is not accuracy.")

    # ---------------------------------------------------------------- 19
    s = d.bullets_slide(
        "Limitations", "What today's AI genuinely cannot do", [
            ("No real understanding", "It matches patterns in language. It does "
                                      "not know what the words mean."),
            ("No common sense", "It can pass a professional exam and then fail a "
                                "question a child would get right."),
            ("Knowledge has a cut-off date", "Unless it is connected to live "
                                             "search, it does not know recent events."),
            ("Weak at exact maths and logic", "It predicts likely text, so it can "
                                              "get simple arithmetic wrong."),
            ("Cannot explain itself", "It usually cannot show why it produced a "
                                      "particular answer."),
            ("No accountability", "A model cannot be responsible. A person always "
                                  "has to be.")],
        sub="Knowing the limits is what separates confident users from careless ones.",
        accent=AMBER, two_col=True)

    # ---------------------------------------------------------------- 20
    s = d.process_slide(
        "Using it well", "A simple four-step habit",
        [("Pick the right task", "Drafting, summarising, explaining - not final "
                                 "decisions."),
         ("Give clear context", "Say the role, the goal, the audience and the "
                                "format."),
         ("Check the output", "Verify every fact, figure, name and quote."),
         ("Keep the human in charge", "A person signs off and owns the result.")],
        accent=[BLUE, TEAL, AMBER, GREEN],
        note="Treat AI as a fast junior colleague: helpful, quick, and always "
             "reviewed before anything goes out.")

    # ---------------------------------------------------------------- 21
    s = d.quote_slide(
        "AI will not replace you. A person using AI well might.",
        "The skill is no longer doing the task. It is directing the tool and "
        "judging the result.",
        kicker="Something to remember")

    # ---------------------------------------------------------------- 22
    s = d.takeaways_slide([
        ("AI learns from data instead of following written rules",
         "That is the one difference that explains everything else."),
        ("Everything in use today is narrow AI",
         "Excellent at one job, useless outside it."),
        ("Predictive AI forecasts, Generative AI creates",
         "Different jobs. Strongest when used together."),
        ("The benefits are speed, cost, scale and consistency",
         "Which is why adoption is moving so fast."),
        ("The risks are real but manageable",
         "Hallucination, bias, data leakage and over-reliance."),
        ("A human stays accountable for the output",
         "Verify, then approve. Always."),
    ], sub="Unit 1 in six lines.")

    # ---------------------------------------------------------------- 23
    d.closing_slide(
        "Thank you",
        "Next up — Unit 2: the platforms, tools and prompt-writing skills that "
        "turn these ideas into daily work.",
        ["Which tasks in your own work are drafting rather than deciding?",
         "Where would a wrong AI answer cost you the most?",
         "What should your team never paste into a public AI tool?"])

    path = f"{outdir}/{OUT}"
    d.save(path)
    return path


if __name__ == "__main__":
    import sys
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
