"""UNIT 2 — Tools and Technologies."""

from deck import (Deck, notes, BLUE, TEAL, AMBER, VIOLET, GREEN, RED)

OUT = "Unit 2 - Tools and Technologies.pptx"


def build(outdir="."):
    d = Deck(2, "Tools and Technologies")

    # ---------------------------------------------------------------- 01
    s = d.title_slide(
        "Tools, Platforms and Prompt Engineering",
        "The Gen AI tools people actually use at work, how to write "
        "instructions that get good results, and how to connect AI to a "
        "process without writing code.",
        ["PGPM · Systems Specialisation", "Unit 2 of 5"],
        chips=["Platforms & Tools", "Prompt Engineering", "APIs", "No-Code"])
    notes(s, "Unit 1 was the theory. This unit is the hands-on one. Tell the "
             "class the goal: by the end they should be able to pick the right "
             "tool and write a prompt that works first time.")

    # ---------------------------------------------------------------- 02
    d.agenda_slide([
        ("The tool landscape", "How to sort dozens of tools into a simple map"),
        ("Six tools up close", "Numerous AI, Jasper, Midjourney, VEO, Odoo, Canva AI"),
        ("Choosing a tool", "Four questions before you buy anything"),
        ("Prompt engineering", "What it is and why the same tool gives some "
                               "people better results"),
        ("The prompt recipe", "Five parts that turn a vague ask into a clear one"),
        ("Types of prompts", "Zero-shot, few-shot, role, chain-of-thought and more"),
        ("APIs explained simply", "How AI gets built into real products"),
        ("No-code automation", "Building useful things with zero programming"),
    ], note="Half theory, half practical. The prompt section is the part you "
            "will use tomorrow.")

    # ---------------------------------------------------------------- 03
    d.section_slide(1, "Platforms and Tools",
                    "There are thousands. You only need a map.", BLUE)

    # ---------------------------------------------------------------- 04
    s = d.cards_slide(
        "The map", "Sort AI tools by the job they do",
        [("Text and writing", "Emails, blogs, reports, summaries.\nChatGPT, Claude, "
                              "Jasper, Copilot."),
         ("Images and design", "Posters, mockups, social posts.\nMidjourney, "
                               "Canva AI, Firefly."),
         ("Video", "Clips, ads, explainers, edits.\nVEO, Runway, Synthesia."),
         ("Audio and voice", "Voiceovers, music, dubbing.\nElevenLabs, Suno, "
                             "Descript."),
         ("Data and spreadsheets", "Clean, classify, analyse rows.\nNumerous AI, "
                                   "Copilot in Excel."),
         ("Business systems", "Run operations with AI built in.\nOdoo, Salesforce, "
                              "SAP.")],
        sub="Five of these create content. The last one runs the business - and "
            "that difference matters.",
        cols=3, accent=BLUE, tinted=True)
    notes(s, "Point out the last card deliberately. The first five are tools you "
             "open to make something. A business system like Odoo is where the "
             "work is actually recorded and run - AI sits inside it.")

    # ---------------------------------------------------------------- 05
    s = d.cards_slide(
        "Tools up close — part 1", "Writing, data and images",
        [("Numerous AI", "Brings AI into Excel and Google Sheets. Drag a formula "
                         "down and it categorises products, cleans messy text or "
                         "writes descriptions for every row.\n\nBest for: bulk "
                         "work on lists and data."),
         ("Jasper", "A writing platform built for marketing teams. Keeps a brand "
                    "voice, uses templates for ads and emails, and works across "
                    "a whole campaign.\n\nBest for: consistent marketing copy at "
                    "volume."),
         ("Midjourney", "Turns a written description into high-quality images. "
                        "Strong on style, mood and art direction.\n\nBest for: "
                        "concepts, moodboards and visuals without a photoshoot.")],
        cols=3, accent=TEAL)

    # ---------------------------------------------------------------- 06
    s = d.cards_slide(
        "Tools up close — part 2", "Video, business systems and everyday design",
        [("VEO", "Google's video generation model. Describe a scene and it "
                 "produces short video with motion, camera movement and matching "
                 "sound.\n\nBest for: quick concept videos and social clips."),
         ("Odoo", "An open-source ERP that runs finance, sales, HR, inventory and "
                  "projects in one place - now with AI features built into those "
                  "modules.\n\nBest for: running the business, not making "
                  "content."),
         ("Canva AI", "AI built into a design tool everyone can already use. "
                      "Generates layouts, rewrites text, removes backgrounds and "
                      "resizes for every channel.\n\nBest for: non-designers who "
                      "need presentable output fast.")],
        cols=3, accent=VIOLET)
    notes(s, "Odoo is the odd one out and worth flagging to the class. The other "
             "five are content tools. Odoo is the system of record - the AI is "
             "embedded in the workflow rather than being the product itself. The "
             "next slide covers it properly.")

    # ---------------------------------------------------------------- 07
    s = d.split_slide(
        "Odoo in detail", "What is Odoo?",
        "An open-source ERP platform that manages business operations - finance, "
        "HR, sales, inventory, CRM and projects - from one system, with AI "
        "features built in.",
        ["ERP means Enterprise Resource Planning",
         "One shared system instead of separate tools per team",
         "Modules: accounting, sales and CRM, HR, inventory",
         "Also manufacturing, projects, e-commerce and point of sale",
         "You install only the modules you need",
         "Used by startups, SMEs and large enterprises alike"],
        [("Why it is in this syllabus", "Recent versions add AI: generated emails "
                                        "and documents, chat assistants and "
                                        "automatic document classification."),
         ("AI inside the workflow", "AI-powered reports and business insights, "
                                    "plus workflow automation and AI agents."),
         ("The wider lesson", "This is where AI is heading - not a separate tool "
                              "you visit, but a feature inside the system you "
                              "already use.")],
        accent=VIOLET)
    notes(s, "If asked for a one-line answer: Odoo is an AI-enabled open-source "
             "ERP used to manage finance, HR, sales, inventory, CRM and projects "
             "from a single platform. Modular - install only what you need.")

    # ---------------------------------------------------------------- 08
    s = d.table_slide(
        "Quick reference", "One table, six tools",
        ["Tool", "What it makes", "Use it for", "Watch out for"],
        [["Numerous AI", "Spreadsheet output", "Sorting and cleaning thousands "
                                               "of rows", "Sending private data "
                                                          "to the cloud"],
         ["Jasper", "Marketing text", "Campaigns that must sound the same",
          "Generic copy if the brief is thin"],
         ["Midjourney", "Images", "Concepts, moodboards, visuals",
          "Rights and usage for commercial work"],
         ["VEO", "Short video", "Ideas, pitches, social clips",
          "Small errors in faces and text"],
         ["Odoo", "Business records and workflows", "Running finance, HR, sales "
                                                    "and inventory with AI help",
          "Setup effort and clean data"],
         ["Canva AI", "Ready-made designs", "Fast, presentable everyday design",
          "Everything can start to look alike"]],
        widths=[1.25, 1.5, 2.2, 1.9], accent=BLUE,
        note="Five of these help you produce something. Odoo is where the business "
             "actually runs.")

    # ---------------------------------------------------------------- 08
    s = d.process_slide(
        "Choosing well", "Four questions before you buy any AI tool",
        [("What is the job?", "Name the task and how often it happens. No task, "
                              "no tool."),
         ("Where does data go?", "Is company data stored, or used for training? "
                                 "Read the terms."),
         ("Does it fit our tools?", "It must work inside email, Office or the CRM "
                                    "you already use."),
         ("What does it save?", "Hours saved per month against the licence cost.")],
        accent=[BLUE, TEAL, AMBER, GREEN],
        note="A tool nobody opens after week two is not a saving. Adoption is "
             "the real test.")

    # ---------------------------------------------------------------- 09
    d.section_slide(2, "Prompt Engineering",
                    "Same tool, very different results. Here is why.", AMBER)

    # ---------------------------------------------------------------- 10
    s = d.split_slide(
        "Definition", "Prompt engineering",
        "The skill of writing instructions that get the output you actually "
        "wanted, first time.",
        ["A prompt is simply the instruction you give the AI",
         "The model cannot read your mind or your context",
         "Vague in, vague out - every single time",
         "Small changes in wording change the result a lot",
         "It is a communication skill, not a coding skill"],
        [("Not a technical job", "If you can brief a new intern clearly, you can "
                                 "write a good prompt."),
         ("It is now a work skill", "Job ads list it. It saves hours every week "
                                    "in any role."),
         ("It is trial and error", "Nobody gets it perfect first time. Good "
                                   "prompts get edited.")],
        accent=AMBER)

    # ---------------------------------------------------------------- 11
    s = d.compare_slide(
        "Why it matters", "The same tool, two very different results",
        {"label": "Weak prompt", "headline": "\"Write about our new product.\"",
         "points": ["No audience, so it guesses",
                    "No length, so it rambles",
                    "No tone, so it sounds generic",
                    "No facts, so it may invent them",
                    "Result: you rewrite the whole thing"]},
        {"label": "Strong prompt", "headline": "\"You are a B2B copywriter. Write "
                                               "a 120-word LinkedIn post...\"",
         "points": ["Audience named: operations managers",
                    "Length fixed: 120 words",
                    "Tone set: clear, no hype",
                    "Facts supplied: three key features",
                    "Result: usable after a light edit"]},
        sub="The model did not get smarter. The instruction did.",
        lacc=RED, racc=GREEN,
        verdict="Every minute spent on the prompt saves ten minutes of rewriting.")

    # ---------------------------------------------------------------- 12
    s = d.cards_slide(
        "The recipe", "Five parts of a good prompt",
        [("1. Role", "Tell it who to be.\n\"You are a financial analyst.\""),
         ("2. Task", "Say exactly what to do.\n\"Summarise this report.\""),
         ("3. Context", "Give the background.\n\"For the board, who are "
                        "non-technical.\""),
         ("4. Format", "Describe the shape.\n\"Five bullets, under 15 words each.\""),
         ("5. Constraints", "Set the limits.\n\"No jargon. Use only the data "
                            "given.\"")],
        sub="Miss a part and the model fills the gap with a guess.",
        cols=5, accent=AMBER, tinted=True)
    notes(s, "Live demo idea: run the same request with and without role and "
             "format. The difference is obvious and the class remembers it.")

    # ---------------------------------------------------------------- 13
    s = d.process_slide(
        "The process", "How prompting actually works in practice",
        [("Define", "Be clear on what a good answer looks like before you type."),
         ("Draft", "Write the first prompt using the five-part recipe."),
         ("Test", "Run it. Compare the output against what you wanted."),
         ("Refine", "Fix the weakest part. Change one thing at a time."),
         ("Reuse", "Save the winning prompt so the whole team can use it.")],
        sub="Prompting is a short loop, not a single shot.",
        accent=[BLUE, TEAL, AMBER, VIOLET, GREEN],
        note="The last step is the one companies skip. A shared prompt library "
             "is where the real time saving comes from.")

    # ---------------------------------------------------------------- 14
    s = d.cards_slide(
        "Types of prompts", "Six patterns worth knowing",
        [("Zero-shot", "Just ask, no examples. Fast, fine for simple tasks."),
         ("Few-shot", "Give two or three examples first. Best way to lock a "
                      "format or style."),
         ("Role prompting", "Assign an expert identity so the tone and depth "
                            "match the job."),
         ("Chain of thought", "Ask it to work through the steps. Better for "
                              "reasoning and analysis."),
         ("Iterative", "Start broad, then refine over several turns like a real "
                       "conversation."),
         ("Template", "A fill-in-the-blank prompt saved for a repeated task.")],
        sub="Match the pattern to the difficulty of the task.",
        cols=3, accent=VIOLET)

    # ---------------------------------------------------------------- 15
    s = d.quadrant_slide(
        "Where it is used", "Prompting at work, by function",
        [("Everyday office work", ["Summarise long email threads and reports",
                                   "Turn rough notes into clean minutes",
                                   "Draft replies you only need to check"], BLUE),
         ("Marketing and sales", ["Campaign copy in one consistent voice",
                                  "Personalised outreach at scale",
                                  "Product descriptions for whole catalogues"], TEAL),
         ("Analysis and reporting", ["Explain what a dataset is showing",
                                     "Draft the commentary around the numbers",
                                     "Convert findings into a board summary"], AMBER),
         ("Learning and support", ["Explain a hard topic at a simpler level",
                                   "Build training material and quizzes",
                                   "Answer routine internal questions"], VIOLET)],
        sub="The same skill pays off in every department.",
        accent=BLUE)

    # ---------------------------------------------------------------- 16
    s = d.compare_slide(
        "Habits", "Do this, not that",
        {"label": "Do", "headline": "Habits of people who get good output",
         "points": ["Say who the audience is",
                    "Give the facts it should use",
                    "Ask for a specific format and length",
                    "Show an example of good work",
                    "Ask for two options and pick one"]},
        {"label": "Avoid", "headline": "Habits that waste your time",
         "points": ["One-line prompts with no context",
                    "Pasting confidential data into public tools",
                    "Changing five things at once when refining",
                    "Trusting names, numbers and quotes blindly",
                    "Sending output out without reading it"]},
        lacc=GREEN, racc=RED,
        verdict="Brief it like a colleague, then review it like an editor.")

    # ---------------------------------------------------------------- 17
    d.section_slide(3, "APIs and No-Code Tools",
                    "Moving from chatting with AI to building with it.", TEAL)

    # ---------------------------------------------------------------- 18
    s = d.split_slide(
        "Plain English", "What is an API?",
        "A doorway that lets one piece of software ask another to do a job.",
        ["Your app sends a request: here is the text, summarise it",
         "The AI service sends an answer back in seconds",
         "No screen, no chat window - software talking to software",
         "You pay for what you use, usually per volume of text",
         "This is how AI gets built into real products"],
        [("Chat window", "One person, one question, one answer. Manual."),
         ("API", "Thousands of requests a day, automatically, inside your own "
                 "system."),
         ("Why care?", "Chat helps a person. An API changes a whole process.")],
        accent=TEAL)
    notes(s, "Analogy: ordering food. In the chat window you walk to the counter "
             "yourself. With an API your system places the order automatically, "
             "every time it is needed.")

    # ---------------------------------------------------------------- 19
    s = d.compare_slide(
        "Two routes", "Build with code, or build with no code",
        {"label": "API route", "headline": "A developer connects AI to your system",
         "points": ["Full control over behaviour and data",
                    "Fits neatly into existing products",
                    "Handles very large volumes",
                    "Needs developers and testing time",
                    "Best for customer-facing features"]},
        {"label": "No-code route", "headline": "A business user assembles it visually",
         "points": ["Drag, drop and connect - no programming",
                    "Working version in hours, not months",
                    "Anyone in the team can maintain it",
                    "Less flexible for unusual cases",
                    "Best for internal workflows"]},
        sub="Most organisations use both: no-code to prove the idea, API to scale it.",
        lacc=BLUE, racc=GREEN,
        verdict="Start no-code to prove it works. Move to an API when it matters.")

    # ---------------------------------------------------------------- 20
    s = d.cards_slide(
        "No-code toolkit", "What people build with, without programming",
        [("Zapier / Make", "Connect apps with triggers. 'When a form is "
                           "submitted, summarise it and post to the team chat.'"),
         ("Custom GPTs", "Upload your own documents and get an assistant that "
                         "answers only from them."),
         ("Microsoft Copilot Studio", "Build assistants inside Teams and Office "
                                      "using company data."),
         ("Power Automate", "Automate approvals and document flows across "
                            "Microsoft tools."),
         ("n8n", "Visual workflow builder for more complex, multi-step "
                 "automations."),
         ("Notion / Airtable AI", "AI columns and summaries directly inside the "
                                  "database your team already uses.")],
        sub="These are the tools behind most 'we automated it' stories.",
        cols=3, accent=GREEN)

    # ---------------------------------------------------------------- 21
    s = d.process_slide(
        "A worked example", "Automating customer feedback in four steps",
        [("Trigger", "A customer submits the feedback form."),
         ("Send to AI", "A saved prompt asks for topic, sentiment and urgency."),
         ("Route it", "Angry and urgent goes straight to a manager."),
         ("Log and report", "Everything lands in a sheet, summarised weekly.")],
        sub="Built in an afternoon with a form, a no-code tool and one good prompt.",
        accent=[TEAL, BLUE, AMBER, VIOLET],
        note="Note the pattern: a trigger, a prompt, a decision, a record. Most "
             "useful automations look exactly like this.")

    # ---------------------------------------------------------------- 22
    s = d.takeaways_slide([
        ("Choose the tool by the output you need",
         "Text, image, video, audio, data or code. Start there."),
        ("Six tools, six different jobs",
         "Numerous AI for data, Jasper for copy, Midjourney for images, VEO for "
         "video, Odoo for business operations, Canva AI for design."),
        ("Prompting is briefing, not coding",
         "Role, task, context, format, constraints."),
        ("Refine one thing at a time, then save the prompt",
         "A shared prompt library is where teams gain real hours."),
        ("APIs put AI inside your systems",
         "Chat helps one person; an API changes a whole process."),
        ("No-code lets you prove the idea this week",
         "Build small, show value, then scale it properly."),
    ], sub="Unit 2 in six lines.")

    # ---------------------------------------------------------------- 23
    d.closing_slide(
        "Thank you",
        "Next up — Unit 3: putting these tools to work in marketing, finance, "
        "HR and operations, and meeting AI agents.",
        ["Which repeated task in your week could a saved prompt handle?",
         "Which tool would you trial first, and how would you measure it?",
         "What is the one workflow you would automate with no-code?"])

    path = f"{outdir}/{OUT}"
    d.save(path)
    return path


if __name__ == "__main__":
    import sys
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
