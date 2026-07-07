SELECT
	item_style_color "SKC",
	"size" "≥ﬂ¬Î",
	store_house "≤÷ø‚",
	SUM ( available_qty ) "ø…”√ø‚¥Ê"
from main_inventory_list 
GROUP BY
	item_style_color,
	"size",
	store_house 
