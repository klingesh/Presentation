Drop the college logo image in this folder (PNG, JPG or WEBP — a transparent PNG
looks best). The build picks up the first image it finds automatically and places
it on the title, section, quote, closing and footer areas of all five decks.

Rebuild after adding it:

    cd src && for u in 1 2 3 4 5; do python3 unit$u.py ..; done
