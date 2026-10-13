# SI-Baseball-Dice-Roll-Allocator
Python script to allocate dice rolls.

To run: Save your workbook after Excel has calculated, then run 
python allocate_dice_rolls.py YourWorkbook.xlsx 

Add --tabs B1 B2 to do only some tabs.


Future enhancement is to shift to do the combined batting row as well as the pitcher tabs.


The algorithm. It’s a three-level optimization:

Minimise the total of the absolute errors between each column’s total and its target.
Among the best results from step 1, make the related-column totals as close as possible to their targets. I used the groups your sheet already defines in rows 129–131: 1B types, 2B types, hits, outs, GO types (G, G*, DP) and FO types (F, (F), SF, (SF)).
Among those, make the worst single-column error as small as possible.

Solved as an integer program, it takes about 4–25 seconds per grid.

Caveat:  Many  assignments tie for best, and the solver picks one. If you care which rolls get which outcomes (say, keeping your current assignment where it’s still optimal), that can be added as extra rules. I haven’t modelled any such preference.

I have not built in any preference for the way that SI typically assigned hits and outs at this time.  I may do that in the future, if only to pretty up the charts.
