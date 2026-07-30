## logo

Drop the college logo here (PNG, JPG or WEBP). The build looks for
`logo-full.*` for the large placements and `logo-mark.*` for the small footer
crest, falling back to the first image it finds.

## fonts

`fonts/*.fntdata` holds the theme typefaces in the format PowerPoint uses for
embedded fonts, so the decks open with the correct type on a machine that does
not have them installed:

- **Sorts Mill Goudy** — headings, titles and figures
- **Quattrocento Sans** — body copy, labels and captions

Both are SIL Open Font License, which permits embedding. Regenerate them from
TTFs with `npx ttf2eot <font>.ttf <name>.fntdata`.

## rebuilding

    cd src && for u in 1 2 3 4 5; do python3 unit$u.py ..; done
