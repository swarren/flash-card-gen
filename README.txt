My daughter is learning basic mathematics. The school provided a set of "flash
cards" so she could practice. However, they only covered multiplication. I
wanted to cover more operations, e.g. add division into the mix. Hence, this
script.

flash-card-template.odt is a manually created template for what the flash
cards will look like. Each piece of text that varies between the different
cards is assigned a code such as "a0" (indicating item "a" on card "0").
The template is set up to be printed double-sided, flipped along the short
side. The small text at the top of a card shows the answer to the problem on
the other side, so the tester doesn't have to flip the card over to work out
the answer. Yes, the tester gets to cheat and just read the answer!

gen-cards.py uses UNO automation to communicate with LibreOffice. I've used
the exact same technique (and code) with OpenOffice.org in the past too; I
believe the only change needed is to fix PYTHONPATH in gen-cards.sh.

It's actually quite possible to build the document up completely from scratch
using the API. However, I find creating the template much easier to do
interactively. Besides, the OOo/LO docs are hard to parse.

For each set of 24 flash cards, the template is opened, a number of search
and replace operations performed to fill in the cards, and the result exported
as an ODT and PDF document.

You can use a tool like pdfsam to concatenate all the PDF files into a single
large PDF to make printing everything simpler.

As an aside, in the past, I've used more complex versions of this script to
auto-generate calendars, filling in day numbers, day of week, holiday dates,
moon phases, photos, and photo descriptions. Perhaps I'll publish that too.
That script uses a wider variety of "API calls".
