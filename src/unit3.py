"""UNIT 3 — Business Application in AI."""

from deck import (Deck, notes, BLUE, TEAL, AMBER, VIOLET, GREEN, RED)

OUT = "Unit 3 - Business Application in AI.pptx"


def build(outdir="."):
    d = Deck(3, "Business Application in AI")

    # ---------------------------------------------------------------- 01
    s = d.title_slide(
        "AI Across the Business",
        "Where AI creates real value in marketing, finance, HR and operations - "
        "and what AI agents change about how work gets done.",
        ["PGPM · Systems Specialisation", "Unit 3 of 5"],
        chips=["Marketing", "Finance", "HR", "Operations", "AI Agents"])
    notes(s, "This is the unit that connects theory to the workplace. Keep asking "
             "the class: which of these happens in your company today?")

    # ---------------------------------------------------------------- 02
    d.agenda_slide([
        ("Where AI pays off", "The four questions that spot a good use case"),
        ("AI in Marketing", "Targeting, content and campaign decisions"),
        ("AI in Finance", "Forecasting, fraud, reporting and risk"),
        ("AI in HR", "Hiring, onboarding, engagement and retention"),
        ("AI in Operations", "Demand, supply chain, quality and maintenance"),
        ("What is an AI agent?", "The step beyond a chat assistant"),
        ("Types of agents", "Six kinds, from simple rules to teams of agents"),
        ("How an agent works", "The six-step loop behind every agent"),
    ], note="Function by function, then the technology that ties it together.")

    # ---------------------------------------------------------------- 03
    d.section_slide(1, "Where AI Actually Pays Off",
                    "Not every problem needs AI. Some are perfect for it.", BLUE)

    # ---------------------------------------------------------------- 04
    s = d.cards_slide(
        "Choosing use cases", "A good AI use case usually has all four",
        [("It repeats often", "Hundreds or thousands of times a month. Rare tasks "
                              "are not worth automating."),
         ("There is data", "History exists in a system already - not only in "
                           "people's heads."),
         ("Rules are fuzzy", "Judgement is involved, so a simple 'if this then "
                             "that' rule does not work."),
         ("Mistakes are survivable", "A wrong answer can be caught and corrected "
                                     "before it does damage.")],
        sub="Fail any one of these and a spreadsheet or a clear process is the "
            "better answer.",
        cols=4, accent=BLUE, tinted=True)

    # ---------------------------------------------------------------- 05
    s = d.split_slide(
        "The value model", "AI creates value in four ways",
        "Every business case eventually lands in one of these buckets.",
        ["Save time - the work gets done faster",
         "Cut cost - fewer hours for the same output",
         "Grow revenue - better targeting and pricing",
         "Reduce risk - catch problems earlier"],
        [("Save time", "Draft, summarise and sort work that used to be manual."),
         ("Cut cost", "Handle routine volume without adding headcount."),
         ("Grow revenue", "Right offer, right customer, right moment."),
         ("Reduce risk", "Spot fraud, defects and churn before they land.")],
        accent=BLUE)

    # ---------------------------------------------------------------- 06
    d.section_slide(2, "AI in Marketing",
                    "The function where adoption started, and moved fastest.", TEAL)

    # ---------------------------------------------------------------- 07
    s = d.cards_slide(
        "Marketing", "Six things marketing teams use AI for",
        [("Customer segments", "Groups customers by real behaviour instead of "
                               "broad age and location buckets."),
         ("Content at scale", "First drafts of ads, emails, posts and product "
                              "descriptions in every variant."),
         ("Personalisation", "Different message and offer for each segment, "
                             "automatically."),
         ("Send-time and channel", "Predicts when and where each customer is "
                                   "most likely to respond."),
         ("Lead scoring", "Ranks enquiries so sales calls the most promising "
                          "ones first."),
         ("Campaign analysis", "Explains which channel actually drove the result, "
                               "not just the last click.")],
        cols=3, accent=TEAL)

    # ---------------------------------------------------------------- 08
    s = d.stats_slide(
        "Marketing outcomes", "What good looks like",
        [("Days to hours", "Campaign production", "Draft copy and creative "
                                                  "variants are ready the same day "
                                                  "the brief lands."),
         ("1-to-1", "Personalisation", "Every customer can get a relevant message, "
                                       "not one mass email."),
         ("Fewer wasted calls", "Sales efficiency", "Lead scoring puts the best "
                                                    "prospects at the top of the "
                                                    "list.")],
        sub="Same team, more output, better targeting.",
        accent=TEAL,
        note="The risk: everyone's marketing starts to sound the same. Brand voice "
             "and human editing matter more than ever.")

    # ---------------------------------------------------------------- 09
    d.section_slide(3, "AI in Finance",
                    "Where accuracy, speed and audit trails all matter.", GREEN)

    # ---------------------------------------------------------------- 10
    s = d.cards_slide(
        "Finance", "Six high-value finance applications",
        [("Forecasting", "Predicts revenue, cash flow and cost using history, "
                         "seasonality and market signals."),
         ("Fraud detection", "Flags transactions that break the customer's normal "
                             "pattern, in real time."),
         ("Credit risk", "Scores applicants on far more signals than a manual "
                         "review could handle."),
         ("Invoice processing", "Reads invoices, matches them to purchase orders "
                                "and routes exceptions."),
         ("Reporting", "Drafts the narrative around the numbers so analysts "
                       "review instead of write."),
         ("Audit support", "Scans every transaction for anomalies rather than "
                           "checking a sample.")],
        sub="Mostly predictive AI, with Gen AI doing the writing.",
        cols=3, accent=GREEN)

    # ---------------------------------------------------------------- 11
    s = d.compare_slide(
        "A worked example", "Fraud detection: before and after",
        {"label": "Rules only", "headline": "Fixed thresholds written by people",
         "points": ["Blocks anything over a set amount",
                    "Fraudsters learn the limits and stay under them",
                    "Large numbers of false alarms",
                    "Every new pattern needs a new rule",
                    "Genuine customers get blocked"]},
        {"label": "With AI", "headline": "Learns each customer's normal behaviour",
         "points": ["Compares to that customer's own history",
                    "Adapts as new fraud patterns appear",
                    "Far fewer false alarms",
                    "Learns continuously from outcomes",
                    "Still needs a human for the edge cases"]},
        lacc=AMBER, racc=GREEN,
        verdict="AI does not replace the rules. It handles what the rules cannot "
                "describe.")

    # ---------------------------------------------------------------- 12
    d.section_slide(4, "AI in Human Resources",
                    "High value, and the highest need for care.", VIOLET)

    # ---------------------------------------------------------------- 13
    s = d.cards_slide(
        "Human resources", "Six ways HR teams use AI",
        [("Job descriptions", "Consistent, inclusive postings drafted in minutes."),
         ("CV screening", "Shortlists against the skills that actually matter for "
                          "the role."),
         ("Interview support", "Suggests questions and structures notes so "
                               "candidates are compared fairly."),
         ("Onboarding help", "An assistant that answers policy questions on day "
                             "one, at any hour."),
         ("Engagement insight", "Finds the themes in thousands of open survey "
                                "comments."),
         ("Attrition risk", "Flags teams where people are likely to leave, so "
                            "managers can act early.")],
        cols=3, accent=VIOLET)

    # ---------------------------------------------------------------- 14
    s = d.bullets_slide(
        "Handle with care", "HR is the highest-risk place to use AI badly",
        [("Bias is inherited", "A model trained on past hiring will copy past "
                               "hiring bias."),
         ("Decisions must be explainable", "A rejected candidate may be entitled "
                                           "to know the reason."),
         ("Data is deeply personal", "Salary, performance and health data need "
                                     "strict access control."),
         ("Regulators are watching", "Hiring AI is treated as high risk in most "
                                     "new AI laws."),
         ("Screen, do not decide", "AI can shortlist. A person must make the "
                                   "final call."),
         ("Tell people it is in use", "Transparency avoids most of the trust "
                                      "problems later.")],
        sub="Use AI to widen the funnel and save admin time, not to pick people.",
        accent=RED, two_col=True)
    notes(s, "Real cases exist of hiring tools quietly down-ranking women because "
             "past hires were mostly men. The tool was working exactly as trained. "
             "That is the point.")

    # ---------------------------------------------------------------- 15
    d.section_slide(5, "AI in Operations",
                    "The quietest wins, and often the biggest.", AMBER)

    # ---------------------------------------------------------------- 16
    s = d.cards_slide(
        "Operations", "Six operational applications",
        [("Demand forecasting", "Predicts what will sell where, so stock sits in "
                                "the right place."),
         ("Inventory planning", "Balances holding cost against the cost of "
                                "running out."),
         ("Route optimisation", "Plans delivery routes around traffic, weight and "
                                "time windows."),
         ("Predictive maintenance", "Flags a machine that is about to fail, "
                                    "before it stops the line."),
         ("Quality inspection", "Cameras spot defects faster and more "
                                "consistently than tired eyes."),
         ("Service automation", "Handles routine customer questions and escalates "
                                "the rest.")],
        cols=3, accent=AMBER)

    # ---------------------------------------------------------------- 17
    s = d.table_slide(
        "Across the business", "One view of all four functions",
        ["Function", "Typical use", "Main benefit", "Main risk"],
        [["Marketing", "Segments, content, personalisation", "Speed and relevance",
          "Generic, off-brand output"],
         ["Finance", "Forecasting, fraud, reporting", "Accuracy and early warning",
          "Over-trusting a model"],
         ["HR", "Screening, onboarding, retention", "Less admin, wider funnel",
          "Bias and legal exposure"],
         ["Operations", "Demand, routes, maintenance", "Lower cost, less downtime",
          "Bad data breaks the forecast"]],
        widths=[1.3, 2.4, 2.0, 2.0], accent=BLUE,
        note="Different functions, same pattern: AI drafts and predicts, people "
             "decide and own it.")

    # ---------------------------------------------------------------- 18
    d.section_slide(6, "AI Agents", "From answering questions to getting things "
                                    "done.", BLUE)

    # ---------------------------------------------------------------- 19
    s = d.compare_slide(
        "The difference", "An assistant answers. An agent acts.",
        {"label": "Chat assistant", "headline": "You do the work, it helps you think",
         "points": ["Waits for each instruction",
                    "Produces text and nothing else",
                    "Forgets once the chat closes",
                    "Cannot touch other systems",
                    "You carry out every step yourself"]},
        {"label": "AI agent", "headline": "You set the goal, it works towards it",
         "points": ["Given an objective, not a script",
                    "Plans the steps by itself",
                    "Uses tools: email, search, databases",
                    "Checks its own progress and retries",
                    "Reports back when the goal is met"]},
        sub="Same underlying model. Very different amount of autonomy.",
        lacc=BLUE, racc=VIOLET,
        verdict="An assistant helps you work. An agent does the work and reports "
                "back.")

    # ---------------------------------------------------------------- 20
    s = d.cards_slide(
        "Types of agents", "Six types, in order of independence",
        [("Simple reflex", "Reacts to what is happening right now. No memory. A "
                           "thermostat."),
         ("Model-based", "Keeps an internal picture of the situation, so it can "
                         "handle what it cannot directly see."),
         ("Goal-based", "Given an outcome, it plans a route to reach it."),
         ("Utility-based", "Weighs trade-offs and picks the best option, not just "
                           "a valid one."),
         ("Learning", "Improves from feedback and results over time."),
         ("Multi-agent", "Several specialist agents working together with a "
                         "coordinator.")],
        sub="More independence means more value, and more supervision needed.",
        cols=3, accent=VIOLET, numbered=True)

    # ---------------------------------------------------------------- 21
    s = d.process_slide(
        "How it works", "The six-step loop inside every agent",
        [("Goal", "A person sets the objective and the limits."),
         ("Plan", "It breaks the goal into ordered steps."),
         ("Act", "It uses tools - search, email, systems, code."),
         ("Observe", "It reads the result of what it just did."),
         ("Adjust", "If a step failed, it retries or changes approach."),
         ("Report", "It hands back the outcome for human sign-off.")],
        accent=[BLUE, TEAL, AMBER, VIOLET, GREEN, RED],
        note="The loop is what makes it an agent. Plan, act, check, adjust - "
             "repeat until done.")

    # ---------------------------------------------------------------- 22
    s = d.split_slide(
        "In practice", "A support agent, end to end",
        "One goal, six steps, one human check at the end.",
        ["Goal: resolve this refund request",
         "Reads the ticket and finds the order",
         "Checks the refund policy and the delivery date",
         "Drafts the reply and prepares the refund",
         "Escalates to a person if the amount is large",
         "Logs everything for the audit trail"],
        [("What people gain", "Routine tickets close in minutes, at any hour."),
         ("What people keep", "The judgement calls, exceptions and the "
                              "relationship."),
         ("What you must add", "Clear limits, an audit log and an easy way to "
                               "escalate.")],
        accent=BLUE)

    # ---------------------------------------------------------------- 23
    s = d.quadrant_slide(
        "Before you deploy", "Four guardrails every agent needs",
        [("Clear limits", ["Define what it may never do alone",
                           "Cap the value it can approve",
                           "Restrict which systems it can reach"], RED),
         ("Human checkpoints", ["Approval before anything irreversible",
                                "A named owner for every agent",
                                "An obvious way to stop it"], AMBER),
         ("Full logging", ["Record every action and decision",
                           "Keep it reviewable and auditable",
                           "Know why it did what it did"], BLUE),
         ("Start small", ["One narrow task, low risk, first",
                          "Measure quality before widening scope",
                          "Expand only when it has earned trust"], GREEN)],
        sub="Autonomy without guardrails is how pilots turn into incidents.",
        accent=RED)

    # ---------------------------------------------------------------- 24
    s = d.takeaways_slide([
        ("Good use cases repeat, have data, need judgement and tolerate error",
         "Check all four before you invest."),
        ("Marketing gains speed and relevance",
         "Segments, content and personalisation at scale."),
        ("Finance gains accuracy and early warning",
         "Forecasting, fraud, invoices and reporting."),
        ("HR gains time but carries the most risk",
         "Screen and support. Never let AI make the final call."),
        ("Operations gains cost and uptime",
         "Demand, routes, maintenance and quality."),
        ("Agents act, so they need guardrails",
         "Limits, human checkpoints, logging, and a small start."),
    ], sub="Unit 3 in six lines.")

    # ---------------------------------------------------------------- 25
    d.closing_slide(
        "Thank you",
        "Next up — Unit 4: privacy, security, regulation and what happens when AI "
        "starts influencing real decisions.",
        ["Which function in your organisation would gain the most, fastest?",
         "Which decision would you never hand to an agent?",
         "What data would you need before any of this could work?"])

    path = f"{outdir}/{OUT}"
    d.save(path)
    return path


if __name__ == "__main__":
    import sys
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
