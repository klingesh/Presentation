"""Budgeting and Savings Habits Among Students — golden red theme."""

import deck as D

D.set_palette("golden")          # must happen before anything is drawn

import embed_fonts                                            # noqa: E402
from deck import Deck, notes                                   # noqa: E402

OUT = "Budgeting and Savings Habits Among Students.pptx"


def build(outdir="."):
    d = Deck(1, "Budgeting and Savings Habits Among Students",
             subject="Personal Finance for Students")
    d.prs.core_properties.title = "Budgeting and Savings Habits Among Students"
    d.prs.core_properties.comments = "Personal finance session for students"

    # ---------------------------------------------------------------- 01
    s = d.title_slide(
        "Budgeting and Savings Habits Among Students",
        "Where student money quietly disappears, a simple way to plan it, and "
        "the small habits that make saving stick — with numbers you can try "
        "this month.",
        chips=["Where Money Goes", "The 50-30-20 Rule", "Savings Ladder",
               "30-Day Plan"])
    notes(s, "Open with a show of hands: who knows roughly what they spent last "
             "week? Most will not. That is the whole reason for this session - "
             "not more income, just visibility.")

    # ---------------------------------------------------------------- 02
    d.agenda_slide([
        ("The student money reality", "Why managing money feels harder than it "
                                      "looks"),
        ("Budgeting vs saving", "Two different habits, often confused"),
        ("Where the money goes", "The four traps that catch almost everyone"),
        ("The leak test", "What a small daily spend really costs in a year"),
        ("The 50-30-20 rule", "One simple split anyone can remember"),
        ("Build your budget", "Five steps, ten minutes, no app needed"),
        ("The savings ladder", "Five levels from first ₹100 to real goals"),
        ("A 30-day starter plan", "What to actually do over the next four weeks"),
    ], note="Half of this is understanding the problem. Half is a plan you can "
            "start on Monday.")

    # ---------------------------------------------------------------- 03
    d.section_slide(1, "The Student Money Reality",
                    "Why this is genuinely harder than people admit.", D.A1)

    # ---------------------------------------------------------------- 04
    s = d.cards_slide(
        "The starting point", "Six reasons student money runs out early",
        [("Income is irregular", "Allowance, part-time work or freelance pay "
                                 "arrive at different times and different "
                                 "amounts."),
         ("First time in charge", "Nobody taught most of us how to plan money. "
                                  "We learn by running out."),
         ("Paying is frictionless", "UPI takes two seconds. There is no pause to "
                                    "think, and no cash left in hand to warn you."),
         ("Group spending", "Plans are made together, so the cost is decided by "
                            "the group, not by your budget."),
         ("No safety net", "One unexpected expense — a phone repair, a trip home "
                           "— can undo a whole month."),
         ("Small feels harmless", "₹40 here and ₹80 there never feels like a "
                                  "decision. Together they are the biggest "
                                  "expense.")],
        sub="None of these are about being careless. They are about how student "
            "money actually works.",
        cols=3, accent=D.A1)

    # ---------------------------------------------------------------- 05
    s = d.compare_slide(
        "Clearing it up", "Budgeting and saving are not the same thing",
        {"label": "Budgeting", "headline": "A plan for money before you spend it",
         "points": ["Decides where each rupee should go",
                    "Happens at the start of the month",
                    "Answers: can I afford this?",
                    "Takes about ten minutes to set up",
                    "Without it, saving rarely happens"]},
        {"label": "Saving", "headline": "Money you deliberately keep aside",
         "points": ["Sets money apart before it can be spent",
                    "Happens on the day money arrives",
                    "Answers: what is this money for?",
                    "Works best when it is automatic",
                    "Grows a buffer and funds real goals"]},
        sub="You need both. A budget makes room; saving fills it.",
        lacc=D.A1, racc=D.A2,
        verdict="Budgeting is the plan. Saving is the result of following it.")

    # ---------------------------------------------------------------- 06
    d.section_slide(2, "Where The Money Goes",
                    "Find the leaks before trying to plug them.", D.A2)

    # ---------------------------------------------------------------- 07
    s = d.quadrant_slide(
        "The four traps", "Where student budgets usually leak",
        [("Food and delivery", ["Ordering in because cooking feels like effort",
                                "Delivery fees and small order charges add up",
                                "Often the single biggest avoidable cost"],
          D.A1),
         ("Subscriptions", ["Free trials that quietly start charging",
                            "Several streaming or music plans at once",
                            "Money leaves without a decision each month"], D.A2),
         ("Sales and impulse buys", ["Discounts create urgency that feels real",
                                     "Saved cards make checkout too easy",
                                     "You buy the offer, not the item"], D.A3),
         ("Outings and peer pressure", ["Saying no feels socially expensive",
                                        "One outing can equal a week's budget",
                                        "Nobody discusses what they can afford"],
          D.A1)],
        sub="Notice that none of these are big single purchases. They are "
            "repeated small ones.",
        accent=D.A2)
    notes(s, "Ask the class which of the four is their own biggest leak. Almost "
             "everyone can name it instantly - which proves the problem is not "
             "knowledge, it is tracking.")

    # ---------------------------------------------------------------- 08
    s = d.stats_slide(
        "The leak test", "What one small daily habit really costs",
        [("₹1,500", "One snack a day, for a month",
          "₹50 a day feels like nothing. Over a month it is often more than a "
          "phone recharge and a subscription combined."),
         ("₹18,000", "The same habit, over a year",
          "That is a laptop upgrade, a certification course, or a serious start "
          "to an emergency buffer."),
         ("₹0", "What it costs to fix it",
          "Nothing. Cutting it from daily to twice a week keeps most of the "
          "money and none of the misery.")],
        sub="Figures are illustrative, but the arithmetic is the point: small "
            "and daily beats large and rare.",
        accent=D.A1,
        note="Run this on your own habit. Daily amount × 30, then × 12. The "
             "number is usually uncomfortable.")

    # ---------------------------------------------------------------- 09
    s = d.bullets_slide(
        "The honest reasons", "Why overspending happens to sensible people",
        [("Small amounts skip the decision", "We judge ₹60 as 'nothing' instead "
                                             "of comparing it to the month."),
         ("Digital payment removes the pause", "Handing over cash makes you feel "
                                               "the spend. A tap does not."),
         ("Offers manufacture urgency", "A deadline makes a want feel like a "
                                        "need."),
         ("Spending is social", "It is easier to overspend with friends than "
                                "alone."),
         ("Boredom costs money", "Scrolling and shopping fill the same empty "
                                 "twenty minutes."),
         ("No record, no reality", "If nothing is written down, the month has no "
                                   "scoreboard.")],
        sub="Understanding the cause matters, because willpower alone is a weak "
            "plan.",
        lead="Overspending is rarely a discipline problem. It is usually a "
             "visibility problem.",
        accent=D.A3, two_col=True)

    # ---------------------------------------------------------------- 10
    d.section_slide(3, "How To Budget Simply",
                    "One rule, five steps, ten minutes a month.", D.A1)

    # ---------------------------------------------------------------- 11
    s = d.cards_slide(
        "The rule", "The 50-30-20 split",
        [("50% — Needs", "Things you cannot skip.\n\nHostel or rent, mess food, "
                         "travel, data, books and course material.", D.A1),
         ("30% — Wants", "Things that make life enjoyable.\n\nEating out, "
                         "subscriptions, outings, clothes, games.", D.A3),
         ("20% — Savings", "Money set aside first.\n\nEmergency buffer, goals, "
                           "and anything you are building towards.", D.A2)],
        sub="On ₹10,000 a month that is ₹5,000 for needs, ₹3,000 for wants and "
            "₹2,000 saved.",
        cols=3, accent=D.A1, tinted=True)
    notes(s, "Stress that the percentages are a starting frame, not a law. If "
             "rent alone is 60%, the split shifts - but the habit of naming "
             "three buckets still works.")

    # ---------------------------------------------------------------- 12
    s = d.process_slide(
        "The method", "Build your budget in five steps",
        [("Know your income", "Add every rupee that comes in each month, from "
                              "every source."),
         ("List fixed needs", "Rent, mess, travel, data. These are decided "
                              "before anything else."),
         ("Save first, not last", "Move your savings amount out on day one, "
                                  "before you can spend it."),
         ("Give wants a limit", "A number, not a feeling. When it is gone, it is "
                                "gone."),
         ("Review weekly", "Five minutes every Sunday. Adjust early instead of "
                           "discovering late.")],
        sub="No app required. A page in a notebook is enough to start.",
        accent=[D.A1, D.A2, D.A3, D.A1, D.A2],
        note="Step three is the one that changes outcomes. Saving what is left "
             "over almost always means saving nothing.")

    # ---------------------------------------------------------------- 13
    s = d.table_slide(
        "A worked example", "One month on ₹10,000",
        ["Category", "Type", "Planned", "How it is decided"],
        [["Mess and food", "Need", "₹3,500", "Fixed before the month starts"],
         ["Travel", "Need", "₹800", "Bus pass or regular commute"],
         ["Phone and data", "Need", "₹400", "Recurring, easy to forget"],
         ["Books and stationery", "Need", "₹300", "Varies by semester"],
         ["Eating out and outings", "Want", "₹1,800", "Hard limit for the month"],
         ["Subscriptions", "Want", "₹400", "Reviewed every few months"],
         ["Shopping and other", "Want", "₹800", "Whatever is left of wants"],
         ["Savings", "Save", "₹2,000", "Moved out on day one"]],
        widths=[2.3, 1.0, 1.1, 3.0], accent=D.A1,
        note="Illustrative figures. Copy the structure, not the numbers — yours "
             "will look different.")

    # ---------------------------------------------------------------- 14
    s = d.cards_slide(
        "Tracking", "Six ways to keep score — pick one you will actually use",
        [("A notebook", "One line per spend. Slow, but you feel every entry, "
                        "which is the point."),
         ("Phone notes", "Always with you. Add the amount the moment you pay."),
         ("Bank and UPI history", "Already recorded for you. Skim it once a week "
                                  "and total the categories."),
         ("A simple spreadsheet", "Four columns: date, item, amount, need or "
                                  "want. Nothing more."),
         ("A budgeting app", "Automatic and tidy, but only helps if you open it."),
         ("The envelope method", "Cash split into labelled envelopes. Brutal, "
                                 "visible and very effective.")],
        sub="The best system is the boring one you still use in week three.",
        cols=3, accent=D.A2)

    # ---------------------------------------------------------------- 15
    d.section_slide(4, "Building The Savings Habit",
                    "Small, automatic and consistent beats large and "
                    "occasional.", D.A2)

    # ---------------------------------------------------------------- 16
    s = d.split_slide(
        "The core idea", "Pay yourself first",
        "Treat savings as a bill you owe yourself, due the day money arrives.",
        ["Decide the amount before the month starts",
         "Move it out on day one, not day thirty",
         "Keep it somewhere slightly inconvenient to reach",
         "Live on what remains, which is easier than it sounds",
         "Raise the amount slightly whenever income rises"],
        [("The wrong order", "Spend the month, then save whatever survives. "
                             "Usually nothing does."),
         ("The right order", "Save first, then spend the rest. The same income "
                             "suddenly produces savings."),
         ("Why it works", "You are not relying on willpower at the end of a "
                          "tiring month.")],
        accent=D.A2)

    # ---------------------------------------------------------------- 17
    s = d.steps_slide(
        "The ladder", "Five levels of a savings habit",
        [("Just start", "Any fixed amount, any frequency. ₹100 a week counts. "
                        "The habit matters more than the sum.", D.A1),
         ("Build a buffer", "Set aside one month of essential expenses. This is "
                            "what stops small emergencies becoming debt.", D.A2),
         ("Name your goals", "Laptop, course fee, trip home. Named money is far "
                             "harder to spend by accident.", D.A3),
         ("Make it automatic", "A standing instruction on allowance day removes "
                               "the decision entirely.", D.A1),
         ("Learn before you grow", "Understand any option properly before you "
                                   "put money into it. Never invest on a tip.",
          D.A2)],
        sub="Most people never get past level one. The gains are at levels three "
            "and four.",
        accent=D.A2)

    # ---------------------------------------------------------------- 18
    s = d.table_slide(
        "Where to keep it", "Common places students park savings",
        ["Option", "How fast you can reach it", "Main drawback", "Suits"],
        [["Cash at home", "Instantly", "Too easy to spend, easy to lose",
          "Very small amounts only"],
         ["Savings bank account", "Same day", "Visible, so tempting to dip into",
          "Your everyday buffer"],
         ["Separate savings account", "Same day", "One more account to manage",
          "Keeping goals away from spending money"],
         ["Recurring deposit", "On maturity", "Penalty if you break it early",
          "Building a fixed monthly habit"],
         ["Fixed deposit", "On maturity", "Locked for the chosen term",
          "Money you know you will not need soon"]],
        widths=[2.1, 1.7, 2.3, 2.2], accent=D.A2,
        note="For classroom learning only, not financial advice. Read the terms "
             "and understand any product before putting money into it.")

    # ---------------------------------------------------------------- 19
    s = d.cards_slide(
        "Making it stick", "Six habits that survive a busy semester",
        [("Save on day one", "Before the month has a chance to spend it for you."),
         ("Use the 24-hour rule", "For any non-essential buy, wait a day. Most "
                                  "urges do not survive it."),
         ("Cap your weak spot", "Set one number for the trap you already know is "
                                "yours."),
         ("Keep one no-spend day", "One day a week where nothing leaves your "
                                   "account. It resets the habit."),
         ("Review for five minutes", "Same time every week. Short and regular "
                                     "beats long and never."),
         ("Keep the goal visible", "A note on your phone screen. Saving is "
                                   "easier when it is for something.")],
        sub="Pick two to start. Six new habits at once is how people quit.",
        cols=3, accent=D.A1)

    # ---------------------------------------------------------------- 20
    s = d.quadrant_slide(
        "Avoid these", "Four mistakes that quietly undo the plan",
        [("Budgeting only in your head", ["Untracked money always feels like more",
                                          "You forget the small spends first",
                                          "Fix: write it down, anywhere"], D.A1),
         ("Saving what is left over", ["At month end there is rarely anything left",
                                       "Savings becomes an accident, not a plan",
                                       "Fix: move it out on day one"], D.A2),
         ("Copying someone else's budget", ["Their rent, family support and "
                                            "course costs are not yours",
                                            "It fails, and you blame yourself",
                                            "Fix: start from your own numbers"],
          D.A3),
         ("Quitting after a bad month", ["One overspent month is data, not "
                                         "failure",
                                         "Most people stop here permanently",
                                         "Fix: reset next month and continue"],
          D.A1)],
        sub="Every one of these is a process problem with a simple fix.",
        accent=D.A1)

    # ---------------------------------------------------------------- 21
    s = d.quote_slide(
        "It is not about how much you earn. It is about how much you keep.",
        "A student who saves ₹500 a month has a habit that will still be working "
        "on a salary twenty times larger.",
        kicker="Worth remembering")

    # ---------------------------------------------------------------- 22
    s = d.process_slide(
        "Start now", "Your first 30 days",
        [("Week 1 — Watch", "Change nothing. Just write down every rupee you "
                            "spend."),
         ("Week 2 — Find one leak", "Total your categories. Pick the single "
                                    "biggest avoidable one."),
         ("Week 3 — Save first", "Decide your amount and move it out the day "
                                 "money arrives."),
         ("Week 4 — Review and repeat", "Compare plan against reality, adjust "
                                        "once, and run it again.")],
        sub="Four weeks is enough to see your own pattern and prove the habit "
            "works.",
        accent=[D.A1, D.A2, D.A3, D.A1],
        note="Week one is the most important and the easiest to skip. You cannot "
             "fix what you have not measured.")

    # ---------------------------------------------------------------- 23
    s = d.takeaways_slide([
        ("Overspending is a visibility problem, not a willpower problem",
         "Write it down and most of the leak fixes itself."),
        ("Budgeting plans money; saving keeps it",
         "You need both, and the budget has to come first."),
        ("Small repeated spends cost the most",
         "₹50 a day is ₹18,000 a year. Do the arithmetic on your own habit."),
        ("50-30-20 is enough structure to start",
         "Needs, wants, savings. Adjust the percentages to your reality."),
        ("Pay yourself first, on day one",
         "Saving what is left over almost always means saving nothing."),
        ("Start with two habits, not six",
         "A boring system you still use in week three beats a perfect one you "
         "abandon."),
    ], sub="The whole session in six lines.", accent=D.A1)

    # ---------------------------------------------------------------- 24
    d.closing_slide(
        "Thank you",
        "Budgeting is not about spending less on everything. It is about "
        "deciding, in advance, what your money is for.",
        ["What is your single biggest avoidable spend this month?",
         "What amount could you move out on day one, starting now?",
         "Which one goal would make saving feel worth it?"])

    path = f"{outdir}/{OUT}"
    d.save(path)
    embed_fonts.embed(path)      # travel with the theme fonts
    return path


if __name__ == "__main__":
    import sys
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
