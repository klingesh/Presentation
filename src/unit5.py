"""UNIT 5 — Strategy, innovation and Leadership."""

from deck import (Deck, notes, BLUE, TEAL, AMBER, VIOLET, GREEN, RED)

OUT = "Unit 5 - Strategy, Innovation and Leadership.pptx"


def build(outdir="."):
    d = Deck(5, "Strategy, Innovation and Leadership")

    # ---------------------------------------------------------------- 01
    s = d.title_slide(
        "Strategy, Innovation and Leadership",
        "How Gen AI changes competition in a digital world, how to build real "
        "business solutions with it, and how leaders turn ordinary teams into "
        "highly effective ones.",
        ["Unit 5 of 5"],
        chips=["Digital World", "Business Solutions", "Leadership", "Maturity"])
    notes(s, "This is the closing unit. Shift the framing from 'what is AI' to "
             "'what do we do about it as managers'.")

    # ---------------------------------------------------------------- 02
    d.agenda_slide([
        ("Gen AI in the digital world", "What actually changed for business"),
        ("Where advantage now comes from", "Why the model itself is not the moat"),
        ("Strategic choices", "Adopt, embed or transform"),
        ("Building a solution", "Six steps from idea to something in production"),
        ("Build, buy or blend", "Choosing the right route"),
        ("Measuring value", "Proving it worked, in business terms"),
        ("Leading the change", "Why most failures are cultural, not technical"),
        ("Five levels of an AI-enabled employee", "From curious to force multiplier"),
    ], note="Strategy first, then execution, then people. People is the part "
            "most organisations get wrong.")

    # ---------------------------------------------------------------- 03
    d.section_slide(1, "Gen AI in the Digital World",
                    "What genuinely changed, and what did not.", BLUE)

    # ---------------------------------------------------------------- 04
    s = d.cards_slide(
        "The shift", "Four things Gen AI changed for business",
        [("Content became cheap", "Producing text, images and video is no longer "
                                  "the bottleneck. Judgement is."),
         ("Software learned language", "Anyone can now instruct a system in plain "
                                      "words instead of through a form."),
         ("Expertise spread out", "A junior with good tools can produce work that "
                                  "used to need a specialist."),
         ("Speed became the norm", "Ideas can be prototyped in days, so slow "
                                   "competitors fall behind faster.")],
        sub="The technology is impressive. The change in expectations is what "
            "matters.",
        cols=4, accent=BLUE, tinted=True)

    # ---------------------------------------------------------------- 05
    s = d.compare_slide(
        "What did not change", "Fundamentals still decide who wins",
        {"label": "What AI changed", "headline": "The cost and speed of doing work",
         "points": ["Drafting, summarising and analysis are fast",
                    "Language is now a valid interface",
                    "Small teams can produce a lot",
                    "Prototypes are cheap to test",
                    "Personalisation is affordable at scale"]},
        {"label": "What stayed the same", "headline": "The things that make a "
                                                     "business work",
         "points": ["Customers still need a real problem solved",
                    "Trust still takes years to build",
                    "Bad processes get faster, not better",
                    "Someone still has to decide and own it",
                    "Execution still beats intention"]},
        lacc=BLUE, racc=TEAL,
        verdict="AI amplifies whatever you already are. Strong operations get "
                "stronger; weak ones get faster at being weak.")

    # ---------------------------------------------------------------- 06
    s = d.cards_slide(
        "Competitive advantage", "The model is not the moat",
        [("Your data", "Competitors can rent the same model. They cannot get your "
                       "customer and operational history."),
         ("Your workflow", "Advantage comes from AI embedded in a process, not "
                           "from a licence."),
         ("Your people's skill", "Teams who know how to direct AI get far more "
                                 "from the same tool."),
         ("Your trust", "Customers share more with organisations they trust, which "
                        "makes your data better still.")],
        sub="Everyone has access to the same models. That is exactly why they are "
            "not the differentiator.",
        cols=4, accent=TEAL)

    # ---------------------------------------------------------------- 07
    s = d.steps_slide(
        "Strategic choices", "Three levels of ambition",
        [("Adopt", "Give teams approved tools and training. Quick wins, low risk, "
                   "easy to copy.", BLUE),
         ("Embed", "Build AI into core workflows and products. Slower, and much "
                   "harder to copy.", TEAL),
         ("Transform", "Change what the business sells or how it is priced. "
                       "Highest risk and highest return.", VIOLET)],
        sub="Most organisations should be doing all three at once, at different "
            "speeds.",
        accent=BLUE)

    # ---------------------------------------------------------------- 08
    d.section_slide(2, "Developing AI-Based Business Solutions",
                    "From an interesting idea to something that runs.", GREEN)

    # ---------------------------------------------------------------- 09
    s = d.process_slide(
        "The build path", "Six steps from problem to production",
        [("Frame the problem", "Start from a business pain, never from the "
                               "technology."),
         ("Check the data", "Does the data exist, and is it good enough?"),
         ("Prototype fast", "Two weeks, no-code, one narrow slice of the problem."),
         ("Measure honestly", "Compare against how the work is done today."),
         ("Deploy with guardrails", "Human review, logging, clear limits, named "
                                    "owner."),
         ("Scale and monitor", "Widen the scope only once quality holds up.")],
        accent=[BLUE, TEAL, AMBER, VIOLET, GREEN, RED],
        note="Step one is where projects are won or lost. 'We should use AI' is "
             "not a problem statement.")

    # ---------------------------------------------------------------- 10
    s = d.split_slide(
        "Framing", "Write the problem before the solution",
        "A good AI project can be described in one sentence, with a number in it.",
        ["Who has the problem, and how often?",
         "How is it handled today, and what does that cost?",
         "What does success look like as a number?",
         "What happens if the AI gets it wrong?",
         "Who owns the outcome once it ships?"],
        [("Weak framing", "\"Let's add AI to customer service.\""),
         ("Strong framing", "\"Cut first-response time on the 4,000 routine "
                            "monthly tickets from 6 hours to 15 minutes.\""),
         ("Why it matters", "The second one can be built, measured and defended in "
                            "a budget meeting.")],
        accent=GREEN)

    # ---------------------------------------------------------------- 11
    s = d.table_slide(
        "Route to market", "Build, buy or blend",
        ["Route", "What it means", "Best when", "Watch out for"],
        [["Buy", "Use a ready-made AI product", "The need is common and speed "
                                               "matters", "Weak fit and ongoing "
                                                          "licence cost"],
         ["Blend", "Configure a platform with your own data", "You need your "
                                                             "context but not "
                                                             "custom code",
          "Lock-in to one vendor"],
         ["Build", "Develop on top of model APIs", "The workflow is your "
                                                  "differentiator", "Cost, talent "
                                                                    "and "
                                                                    "maintenance"]],
        widths=[1.1, 2.3, 2.5, 2.3], accent=VIOLET,
        note="Buy for common problems. Build only where it is genuinely your "
             "competitive edge.")

    # ---------------------------------------------------------------- 12
    s = d.stats_slide(
        "Measuring value", "Three numbers a sponsor will ask for",
        [("Time", "Hours returned", "Hours saved per month across the team, valued "
                                    "at a realistic hourly cost."),
         ("Quality", "Error and rework rate", "Fewer mistakes, fewer escalations, "
                                              "fewer things done twice."),
         ("Outcome", "Business result", "Revenue, retention, cost or risk moving "
                                       "in the right direction.")],
        sub="If you cannot measure it, you cannot defend the budget next year.",
        accent=GREEN,
        note="Always measure against a baseline captured before you started. "
             "Without a baseline, every result is an opinion.")

    # ---------------------------------------------------------------- 13
    s = d.quadrant_slide(
        "Why projects fail", "Four failure patterns, and the fix",
        [("Technology-first", ["Started from the tool, not the problem",
                               "Fix: write the problem statement first"], RED),
         ("Pilot purgatory", ["Endless demos, nothing in production",
                              "Fix: set a go or stop date up front"], AMBER),
         ("No owner", ["IT built it, nobody runs it",
                       "Fix: name a business owner on day one"], VIOLET),
         ("No adoption", ["It works, but nobody uses it",
                          "Fix: involve users while building, not after"], BLUE)],
        sub="Most AI projects fail for management reasons, not technical ones.",
        accent=RED)

    # ---------------------------------------------------------------- 14
    d.section_slide(3, "Leadership and People",
                    "The hardest part is never the technology.", AMBER)

    # ---------------------------------------------------------------- 15
    s = d.cards_slide(
        "Leading the change", "Six things leaders must actually do",
        [("Go first", "Use the tools yourself. Teams copy behaviour, not memos."),
         ("Make it safe", "If people fear replacement, they hide what they "
                          "automate."),
         ("Be clear on limits", "Publish what is allowed, so people stop guessing."),
         ("Fund the learning", "Protected time to practise beats another tool "
                               "licence."),
         ("Reward sharing", "Celebrate the person who shares a working prompt, not "
                            "just the biggest project."),
         ("Redesign the work", "If the process stays the same, you get a faster "
                               "version of yesterday.")],
        sub="Adoption is a leadership behaviour, not a training budget.",
        cols=3, accent=AMBER)
    notes(s, "The fear point is critical. If staff believe efficiency leads to job "
             "cuts, they will quietly not tell you what AI can do. Say out loud "
             "what the productivity gain will be used for.")

    # ---------------------------------------------------------------- 16
    s = d.section_slide(4, "From Employee to Force Multiplier",
                        "Five levels of AI capability in a team.", VIOLET)

    # ---------------------------------------------------------------- 17
    s = d.steps_slide(
        "The maturity ladder", "Five levels to a super-efficient employee",
        [("Curious", "Has tried a chatbot once or twice. No routine, no habit.",
          BLUE),
         ("User", "Uses AI daily for drafting and summarising. Saves a few hours "
                  "a week.", TEAL),
         ("Practitioner", "Writes strong prompts, saves and reuses them, checks "
                          "output properly.", AMBER),
         ("Builder", "Automates whole workflows with no-code tools. Saves the "
                     "team's time, not just their own.", VIOLET),
         ("Multiplier", "Redesigns how the team works and teaches others. Raises "
                        "the whole function.", GREEN)],
        sub="Most people are stuck between level 1 and 2. The value is at 4 and 5.",
        accent=VIOLET)

    # ---------------------------------------------------------------- 18
    s = d.table_slide(
        "Moving up a level", "What it takes at each step",
        ["Level", "Behaviour", "Value created", "What unlocks the next level"],
        [["1 · Curious", "Occasional experiments", "Almost none",
          "A relevant use case in their own job"],
         ["2 · User", "Daily drafting and summarising", "A few hours a week",
          "Learning the five-part prompt recipe"],
         ["3 · Practitioner", "Reusable prompts, verified output", "Consistent "
                                                                  "quality and "
                                                                  "speed",
          "Exposure to no-code automation"],
         ["4 · Builder", "Automates real workflows", "Whole-team time savings",
          "Authority to change how work is done"],
         ["5 · Multiplier", "Redesigns and teaches", "Function-level performance",
          "A role that rewards enabling others"]],
        widths=[1.3, 2.1, 2.0, 2.6], accent=VIOLET,
        note="Notice the pattern: the first three levels are skill. The last two "
             "are permission.")

    # ---------------------------------------------------------------- 19
    s = d.bullets_slide(
        "The skills that travel", "What stays valuable as AI gets better",
        [("Asking the right question", "Framing the problem is the work AI cannot "
                                       "do for you."),
         ("Judging quality", "Knowing when a confident answer is actually wrong."),
         ("Context", "Understanding your customers, your business, your "
                     "constraints."),
         ("Decision-making", "Choosing under uncertainty, and owning the "
                             "consequence."),
         ("Communication", "Persuading people, which is still a human-to-human "
                           "act."),
         ("Learning speed", "The tools will change again. The ability to adapt is "
                            "the real asset.")],
        sub="Notice what is missing: producing the first draft.",
        accent=TEAL, two_col=True)

    # ---------------------------------------------------------------- 20
    s = d.process_slide(
        "Where to start", "A realistic first 90 days",
        [("Days 1-30", "Approve tools, publish a one-page policy, train everyone "
                       "once."),
         ("Days 31-60", "Pick two narrow use cases. Prototype both. Capture a "
                        "baseline."),
         ("Days 61-90", "Keep the one that worked, kill the other, and publish the "
                        "numbers.")],
        sub="Small, fast, measured - then repeat.",
        accent=[BLUE, AMBER, GREEN],
        note="Publishing real numbers - including the failure - is what earns "
             "budget and trust for the next round.")

    # ---------------------------------------------------------------- 21
    s = d.quote_slide(
        "Strategy is not choosing a model. It is choosing which work to redesign.",
        "The tools are available to everyone. What you do with them is not.",
        kicker="The closing thought")

    # ---------------------------------------------------------------- 22
    s = d.takeaways_slide([
        ("Gen AI made content cheap and judgement valuable",
         "The bottleneck moved from producing to deciding."),
        ("The model is not the moat",
         "Your data, workflow, people and trust are."),
        ("Adopt, embed, transform - run all three",
         "Different speeds, different levels of risk."),
        ("Start from a problem with a number in it",
         "Frame it, prototype in weeks, measure against a baseline."),
        ("Most failures are management failures",
         "No owner, no adoption, no decision date."),
        ("The value is at levels four and five",
         "Builders and multipliers, unlocked by permission as much as skill."),
    ], sub="Unit 5 in six lines.")

    # ---------------------------------------------------------------- 23
    d.closing_slide(
        "Thank you",
        "That completes the five units — from what AI is, through tools and "
        "applications, to governance, strategy and leadership.",
        ["Which one workflow would you redesign first, and why that one?",
         "What level are you on the ladder today, and what is your next step?",
         "What will your organisation do with the time AI gives back?"])

    path = f"{outdir}/{OUT}"
    d.save(path)
    return path


if __name__ == "__main__":
    import sys
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
