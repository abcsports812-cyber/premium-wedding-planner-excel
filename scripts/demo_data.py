"""Realistic fictional demo data for Sophia & Daniel's wedding, 15 June 2027."""
import random
random.seed(42)

FIRST = ["Olivia", "Emma", "Ava", "Isabella", "Mia", "Charlotte", "Amelia", "Harper", "Evelyn",
         "Abigail", "Emily", "Elizabeth", "Sofia", "Avery", "Ella", "Scarlett", "Grace", "Chloe",
         "Victoria", "Riley", "James", "William", "Benjamin", "Lucas", "Henry", "Alexander",
         "Michael", "Daniel", "Jacob", "Logan", "Jackson", "Sebastian", "Owen", "Samuel", "Matthew",
         "Joseph", "Levi", "David", "John", "Carter", "Nora", "Hazel", "Aurora", "Savannah",
         "Audrey", "Brooklyn", "Bella", "Claire", "Skylar", "Lucy"]
LAST = ["Bennett", "Whitfield", "Carter", "Mitchell", "Sullivan", "Reyes", "Hughes", "Foster",
        "Coleman", "Barrett", "Griffin", "Simmons", "Warren", "Fitzgerald", "Wallace", "Chambers",
        "Nakamura", "Osei", "Patel", "Kowalski", "Moreau", "Delgado", "Lindqvist", "Okafor",
        "Marsh", "Hartley", "Abernathy", "Quinn", "Pruitt", "Sinclair", "Vance", "Whitaker",
        "Ashworth", "Renner", "Castellano", "Beaumont"]

random.shuffle(FIRST)
random.shuffle(LAST)


def full_names(n):
    names = []
    used = set()
    i = 0
    while len(names) < n:
        f = FIRST[i % len(FIRST)]
        l = LAST[(i * 3 + 7) % len(LAST)]
        nm = f"{f} {l}"
        if nm not in used:
            used.add(nm)
            names.append((f, l))
        i += 1
    return names

# ----------------------------------------------------------------- BUDGET --
BUDGET_ITEMS = [
    ("Venue", "Ceremony & Reception", "Vineyard estate rental — full day", "Silver Oak Estate", 12000, 12000, 6000, 12000, "2027-01-15"),
    ("Venue", "Ceremony & Reception", "Site fee & coordinator", "Silver Oak Estate", 1500, 1500, 750, 1500, "2027-01-15"),
    ("Catering", "Dinner Service", "Plated 3-course dinner (150 guests)", "Harvest Table Catering", 9500, 9750, 3000, 6000, "2027-03-01"),
    ("Catering", "Bar Service", "Open bar, 5 hours", "Harvest Table Catering", 3200, 3350, 1000, 3350, "2027-03-01"),
    ("Photography", "Full Day Coverage", "10-hour wedding photography package", "Golden Hour Photo Co.", 3800, 3800, 1000, 1900, "2027-02-01"),
    ("Videography", "Highlight Film", "Cinematic highlight film + full ceremony", "Lumière Films", 2600, 2600, 800, 800, "2027-02-01"),
    ("Attire", "Wedding Dress", "Designer gown + alterations", "Maison Everly Bridal", 2800, 2950, 500, 2950, "2026-11-01"),
    ("Attire", "Groom Suit", "Tailored suit + accessories", "Sartorial House", 900, 875, 200, 875, "2026-11-15"),
    ("Beauty", "Hair & Makeup", "Bridal party hair + makeup, day of", "Bloom Beauty Studio", 1400, 1450, 300, 300, "2027-05-01"),
    ("Flowers", "Ceremony & Reception Florals", "Bouquets, centerpieces, arch florals", "Wildroot Floral Design", 4200, 4400, 1200, 2200, "2027-04-01"),
    ("Decor", "Reception Styling", "Linens, lounge furniture, lighting", "Ember & Ivy Events", 3100, 3050, 800, 1500, "2027-04-01"),
    ("Entertainment", "DJ & MC", "6-hour DJ package with MC services", "Skyline Sound Co.", 1800, 1800, 400, 900, "2027-03-15"),
    ("Music", "Ceremony Strings", "String quartet, ceremony", "Napa String Quartet", 900, 900, 200, 900, "2027-04-15"),
    ("Transportation", "Guest Shuttles", "Shuttle service, 2 routes", "Valley Transit Co.", 1200, 1150, 0, 1150, "2027-05-20"),
    ("Invitations", "Save-the-Dates", "Printed & mailed save-the-dates", "Paperie & Pine", 650, 620, 0, 620, "2026-09-01"),
    ("Stationery", "Invitation Suite", "Invitations, RSVP cards, menus", "Paperie & Pine", 1200, 1180, 0, 1180, "2027-01-10"),
    ("Cake", "Wedding Cake", "4-tier cake + dessert table", "Sugar & Sage Bakery", 1100, 1150, 300, 1150, "2027-05-15"),
    ("Rings", "Wedding Bands", "Matching platinum wedding bands", "Heirloom & Co. Jewelers", 2200, 2150, 0, 2150, "2027-01-01"),
    ("Rentals", "Tables, Chairs & Tableware", "Chiavari chairs, china, glassware", "Vintage Rentals Co.", 2400, 2500, 600, 1200, "2027-04-20"),
    ("Accommodation", "Wedding Weekend Block", "Hotel room block, 20 rooms", "Napa Ridge Inn", 1600, 1600, 400, 1600, "2027-05-01"),
    ("Gifts", "Wedding Party Gifts", "Gifts for bridesmaids & groomsmen", "Local Artisan Market", 700, 680, 0, 680, "2027-05-25"),
    ("Honeymoon", "Honeymoon Trip", "See Honeymoon Budget tab for detail", "Various", 8000, 6450, 2000, 4450, "2027-06-20"),
    ("Miscellaneous", "Contingency Fund", "Unplanned expenses buffer", "N/A", 1150, 400, 0, 400, "2027-06-01"),
]

# --------------------------------------------------------------- VENDORS ---
VENDORS = [
    ("Silver Oak Estate", "Venue", "Marisol Herrera", "707-555-0142", "events@silveroakestate.com", "silveroakestate.com", 13500, 13500, 6750, "Signed", "Deposit Paid", "2027-01-15", 5),
    ("Harvest Table Catering", "Catering", "Grant Whitfield", "707-555-0198", "grant@harvesttablecatering.com", "harvesttablecatering.com", 13100, 13100, 4000, "Signed", "Deposit Paid", "2027-03-01", 5),
    ("Golden Hour Photo Co.", "Photography", "Ines Callahan", "415-555-0110", "ines@goldenhourphoto.co", "goldenhourphoto.co", 3800, 3800, 1000, "Signed", "Deposit Paid", "2027-02-01", 5),
    ("Lumière Films", "Videography", "Theo Marsh", "415-555-0177", "hello@lumierefilms.com", "lumierefilms.com", 2600, 2600, 800, "Signed", "Deposit Paid", "2027-02-01", 4),
    ("Maison Everly Bridal", "Attire", "Camille Everly", "707-555-0133", "camille@maisoneverly.com", "maisoneverly.com", 2950, 2950, 500, "Signed", "Deposit Paid", "2026-11-01", 5),
    ("Sartorial House", "Attire", "Miles Hartley", "707-555-0166", "miles@sartorialhouse.com", "sartorialhouse.com", 875, 875, 200, "Signed", "Deposit Paid", "2026-11-15", 4),
    ("Bloom Beauty Studio", "Beauty", "Priya Nakamura", "707-555-0189", "priya@bloombeautystudio.com", "bloombeautystudio.com", 1450, 1450, 300, "Signed", "Not Started", "2027-05-01", 5),
    ("Wildroot Floral Design", "Flowers", "Rosalind Abernathy", "707-555-0121", "rosalind@wildrootfloral.com", "wildrootfloral.com", 4400, 4400, 1200, "Signed", "Deposit Paid", "2027-04-01", 5),
    ("Ember & Ivy Events", "Decor", "Julian Reyes", "707-555-0144", "julian@emberandivy.com", "emberandivy.com", 3050, 3050, 800, "Signed", "Deposit Paid", "2027-04-01", 4),
    ("Skyline Sound Co.", "Entertainment", "Devon Okafor", "415-555-0155", "devon@skylinesound.com", "skylinesound.com", 1800, 1800, 400, "Signed", "Deposit Paid", "2027-03-15", 5),
    ("Napa String Quartet", "Music", "Elena Kowalski", "707-555-0161", "elena@napastrings.com", "napastrings.com", 900, 900, 200, "Signed", "Paid in Full", "2027-04-15", 5),
    ("Valley Transit Co.", "Transportation", "Marcus Vance", "707-555-0173", "marcus@valleytransit.com", "valleytransit.com", 1150, 1150, 0, "Sent", "Not Started", "2027-05-20", 4),
    ("Paperie & Pine", "Stationery", "Nora Fitzgerald", "707-555-0128", "nora@paperieandpine.com", "paperieandpine.com", 1800, 1800, 0, "Signed", "Overdue", "2027-01-10", 5),
    ("Sugar & Sage Bakery", "Cake", "Willa Chambers", "707-555-0119", "willa@sugarandsage.com", "sugarandsage.com", 1150, 1150, 300, "Signed", "Deposit Paid", "2027-05-15", 5),
    ("Heirloom & Co. Jewelers", "Rings", "Adrian Pruitt", "707-555-0104", "adrian@heirloomjewelers.com", "heirloomjewelers.com", 2150, 2150, 0, "Signed", "Paid in Full", "2027-01-01", 5),
    ("Vintage Rentals Co.", "Rentals", "Bianca Delgado", "707-555-0187", "bianca@vintagerentalsco.com", "vintagerentalsco.com", 2500, 2500, 600, "Signed", "Deposit Paid", "2027-04-20", 4),
    ("Napa Ridge Inn", "Accommodation", "Colin Sinclair", "707-555-0192", "colin@naparidgeinn.com", "naparidgeinn.com", 1600, 1600, 400, "Signed", "Deposit Paid", "2027-05-01", 4),
    ("Local Artisan Market", "Gifts", "Hazel Renner", "707-555-0158", "hazel@localartisanmarket.com", "localartisanmarket.com", 680, 680, 0, "Not Sent", "Not Started", "2027-05-25", 4),
]

# ------------------------------------------------------------- GUEST LIST --
SIDES = ["Partner 1", "Partner 2", "Both"]
RELATIONSHIPS = ["Immediate Family", "Extended Family", "College Friend", "Work Colleague",
                  "Childhood Friend", "Neighbor", "Family Friend"]
RSVP = ["Attending", "Attending", "Attending", "Pending", "Attending", "Declined", "Tentative"]
MEALS = ["Standard", "Vegetarian", "Vegan", "Gluten-Free", "Kids Meal"]


def build_guest_rows(n=58):
    names = full_names(n)
    rows = []
    for i, (f, l) in enumerate(names):
        side = SIDES[i % 3]
        rel = RELATIONSHIPS[i % len(RELATIONSHIPS)]
        rsvp = RSVP[i % len(RSVP)]
        plus_one = "Yes" if i % 4 == 0 else "No"
        plus_one_name = f"{FIRST[(i*5+3)%len(FIRST)]} {LAST[(i*2+1)%len(LAST)]}" if plus_one == "Yes" else ""
        meal = MEALS[i % len(MEALS)] if rsvp == "Attending" else ""
        children = "Yes" if i % 9 == 0 else "No"
        invitation = "Invitation Sent" if i % 11 != 0 else "Save-the-Date Sent"
        table = (i % 15) + 1 if rsvp == "Attending" else ""
        rows.append({
            "Guest Name": f"{f} {l}",
            "Household / Group": f"{l} Family" if rel == "Immediate Family" or rel == "Extended Family" else f"{f} {l}",
            "Side": side,
            "Relationship": rel,
            "Email": f"{f.lower()}.{l.lower()}@example.com",
            "Phone": f"({200+i%700}) 555-{1000+i:04d}",
            "Invitation Sent": invitation,
            "RSVP Status": rsvp,
            "Plus One": plus_one,
            "Plus One Name": plus_one_name,
            "Meal Preference": meal,
            "Dietary Restrictions": "None" if meal in ("Standard", "") else meal,
            "Children Attending": children,
            "Table Number": table,
            "Notes": "",
        })
    return rows

# ------------------------------------------------------------------ TASKS --
CHECKLIST_TASKS = [
    ("12+ Months Before", "Set overall wedding budget", "Planning", "High", "Both", "2026-05-01", "Completed"),
    ("12+ Months Before", "Create guest list draft", "Planning", "High", "Both", "2026-05-15", "Completed"),
    ("12+ Months Before", "Book venue", "Venue", "High", "Partner 1", "2026-06-01", "Completed"),
    ("12+ Months Before", "Hire wedding planner", "Planning", "Medium", "Both", "2026-06-15", "Completed"),
    ("12+ Months Before", "Research photographers", "Photography", "Medium", "Partner 2", "2026-07-01", "Completed"),
    ("9-12 Months Before", "Book photographer", "Photography", "High", "Partner 2", "2026-08-01", "Completed"),
    ("9-12 Months Before", "Book videographer", "Videography", "Medium", "Partner 2", "2026-08-15", "Completed"),
    ("9-12 Months Before", "Book caterer", "Catering", "High", "Both", "2026-09-01", "Completed"),
    ("9-12 Months Before", "Send save-the-dates", "Stationery", "High", "Both", "2026-09-15", "Completed"),
    ("9-12 Months Before", "Choose wedding party", "Wedding Party", "Medium", "Both", "2026-09-01", "Completed"),
    ("9-12 Months Before", "Book entertainment / DJ", "Entertainment", "Medium", "Partner 1", "2026-10-01", "Completed"),
    ("6-9 Months Before", "Shop for wedding dress", "Attire", "High", "Partner 1", "2026-10-15", "Completed"),
    ("6-9 Months Before", "Shop for suit", "Attire", "Medium", "Partner 2", "2026-10-15", "Completed"),
    ("6-9 Months Before", "Book florist", "Flowers", "High", "Both", "2026-11-01", "Completed"),
    ("6-9 Months Before", "Book hair & makeup artist", "Beauty", "Medium", "Partner 1", "2026-11-15", "Completed"),
    ("6-9 Months Before", "Plan honeymoon", "Honeymoon", "Medium", "Both", "2026-12-01", "Completed"),
    ("6-9 Months Before", "Book accommodation block", "Accommodation", "Low", "Partner 1", "2026-12-01", "Completed"),
    ("3-6 Months Before", "Order invitations", "Stationery", "High", "Both", "2027-01-10", "Completed"),
    ("3-6 Months Before", "Book transportation", "Transportation", "Medium", "Partner 2", "2027-01-20", "Completed"),
    ("3-6 Months Before", "Finalize menu with caterer", "Catering", "High", "Both", "2027-02-01", "Completed"),
    ("3-6 Months Before", "Order wedding cake", "Cake", "Medium", "Partner 1", "2027-02-15", "Completed"),
    ("3-6 Months Before", "Purchase wedding rings", "Rings", "High", "Partner 2", "2027-01-01", "Completed"),
    ("3-6 Months Before", "Book rentals (tables, chairs)", "Rentals", "Medium", "Wedding Planner", "2027-02-20", "In Progress"),
    ("3-6 Months Before", "Schedule dress fittings", "Attire", "Medium", "Partner 1", "2027-03-01", "In Progress"),
    ("1-3 Months Before", "Mail invitations", "Stationery", "High", "Both", "2027-03-15", "In Progress"),
    ("1-3 Months Before", "Finalize seating chart draft", "Seating", "Medium", "Both", "2027-04-01", "Not Started"),
    ("1-3 Months Before", "Confirm all vendor contracts", "Vendors", "High", "Wedding Planner", "2027-04-01", "Not Started"),
    ("1-3 Months Before", "Purchase wedding party gifts", "Gifts", "Low", "Both", "2027-05-01", "Not Started"),
    ("1-3 Months Before", "Apply for marriage license", "Legal", "High", "Both", "2027-05-15", "Not Started"),
    ("1-3 Months Before", "Book hair & makeup trial", "Beauty", "Medium", "Partner 1", "2027-04-15", "Not Started"),
    ("Final Month", "Final dress fitting", "Attire", "High", "Partner 1", "2027-05-20", "Not Started"),
    ("Final Month", "Confirm final guest count with caterer", "Catering", "High", "Both", "2027-05-25", "Not Started"),
    ("Final Month", "Finalize seating chart", "Seating", "High", "Both", "2027-05-28", "Not Started"),
    ("Final Month", "Confirm timeline with vendors", "Vendors", "High", "Wedding Planner", "2027-06-01", "Not Started"),
    ("Final Month", "Pick up rings", "Rings", "Medium", "Partner 2", "2027-06-05", "Not Started"),
    ("Final Month", "Pack for honeymoon", "Honeymoon", "Low", "Both", "2027-06-10", "Not Started"),
    ("Wedding Week", "Rehearsal dinner", "Events", "High", "Both", "2027-06-13", "Not Started"),
    ("Wedding Week", "Deliver final payments to vendors", "Payments", "High", "Wedding Planner", "2027-06-13", "Not Started"),
    ("Wedding Week", "Confirm transportation schedule", "Transportation", "Medium", "Partner 2", "2027-06-14", "Not Started"),
    ("Wedding Week", "Drop off decor items", "Decor", "Medium", "Wedding Planner", "2027-06-14", "Not Started"),
    ("Wedding Day", "Hair & makeup", "Beauty", "High", "Partner 1", "2027-06-15", "Not Started"),
    ("Wedding Day", "First look photos", "Photography", "Medium", "Both", "2027-06-15", "Not Started"),
    ("Wedding Day", "Ceremony", "Events", "High", "Both", "2027-06-15", "Not Started"),
    ("Wedding Day", "Reception", "Events", "High", "Both", "2027-06-15", "Not Started"),
    ("Post-Wedding", "Send thank-you cards", "Stationery", "Medium", "Both", "2027-07-15", "Not Started"),
    ("Post-Wedding", "Return rentals", "Rentals", "Medium", "Wedding Planner", "2027-06-17", "Not Started"),
    ("Post-Wedding", "Order wedding album", "Photography", "Low", "Both", "2027-08-01", "Not Started"),
    ("Post-Wedding", "Change name on documents", "Legal", "Low", "Partner 1", "2027-09-01", "Not Started"),
]

# --------------------------------------------------------------- PAYMENTS --
PAYMENTS = [
    ("Silver Oak Estate", "Venue rental balance", 13500, 6750, 6750, 0, "2027-05-15", "Deposit Paid", "Bank Transfer"),
    ("Harvest Table Catering", "Catering final balance", 13100, 4000, 4000, 5100, "2027-05-20", "Partially Paid", "Bank Transfer"),
    ("Golden Hour Photo Co.", "Photography balance", 3800, 1000, 900, 1900, "2027-05-01", "Partially Paid", "Credit Card"),
    ("Lumière Films", "Videography balance", 2600, 800, 0, 1800, "2027-05-01", "Deposit Paid", "Credit Card"),
    ("Maison Everly Bridal", "Dress final payment", 2950, 500, 2450, 0, "2027-04-01", "Paid in Full", "Credit Card"),
    ("Sartorial House", "Suit final payment", 875, 200, 675, 0, "2027-04-01", "Paid in Full", "Credit Card"),
    ("Wildroot Floral Design", "Floral balance", 4400, 1200, 1000, 2200, "2027-05-25", "Partially Paid", "Bank Transfer"),
    ("Ember & Ivy Events", "Decor balance", 3050, 800, 700, 1550, "2027-05-25", "Partially Paid", "Bank Transfer"),
    ("Skyline Sound Co.", "DJ balance", 1800, 400, 500, 900, "2027-05-15", "Partially Paid", "Digital Wallet"),
    ("Napa String Quartet", "Quartet full payment", 900, 200, 700, 0, "2027-04-15", "Paid in Full", "Cheque"),
    ("Valley Transit Co.", "Shuttle service payment", 1150, 0, 0, 1150, "2027-06-01", "Not Started", "Bank Transfer"),
    ("Paperie & Pine", "Stationery balance", 1800, 0, 1200, 600, "2027-02-01", "Overdue", "Credit Card"),
    ("Sugar & Sage Bakery", "Cake balance", 1150, 300, 850, 0, "2027-05-20", "Paid in Full", "Cash"),
    ("Heirloom & Co. Jewelers", "Rings full payment", 2150, 0, 2150, 0, "2027-01-01", "Paid in Full", "Credit Card"),
    ("Vintage Rentals Co.", "Rentals balance", 2500, 600, 400, 1500, "2027-05-30", "Partially Paid", "Bank Transfer"),
    ("Napa Ridge Inn", "Room block deposit", 1600, 400, 1200, 0, "2027-05-01", "Paid in Full", "Bank Transfer"),
    ("Local Artisan Market", "Gifts payment", 680, 0, 0, 680, "2027-05-25", "Not Started", "Cash"),
]

# ---------------------------------------------------------- WEDDING PARTY --
WEDDING_PARTY = [
    ("Isabella Carter", "Maid of Honor", "Partner 1", "555-0201", "isabella.carter@example.com", "Coordinate bridesmaids, help with dress", "Dusty rose gown", ""),
    ("Grace Sullivan", "Bridesmaid", "Partner 1", "555-0202", "grace.sullivan@example.com", "Bridal shower planning", "Dusty rose gown", ""),
    ("Chloe Mitchell", "Bridesmaid", "Partner 1", "555-0203", "chloe.mitchell@example.com", "Guest book duties", "Dusty rose gown", ""),
    ("Ava Reyes", "Bridesmaid", "Partner 1", "555-0204", "ava.reyes@example.com", "Help with seating chart", "Dusty rose gown", ""),
    ("William Foster", "Best Man", "Partner 2", "555-0205", "william.foster@example.com", "Coordinate groomsmen, toast", "Charcoal suit", ""),
    ("Lucas Griffin", "Groomsman", "Partner 2", "555-0206", "lucas.griffin@example.com", "Bachelor party planning", "Charcoal suit", ""),
    ("Henry Warren", "Groomsman", "Partner 2", "555-0207", "henry.warren@example.com", "Transportation coordination", "Charcoal suit", ""),
    ("Owen Hughes", "Groomsman", "Partner 2", "555-0208", "owen.hughes@example.com", "Ring bearer supervision", "Charcoal suit", ""),
    ("Nora Fitzgerald", "Officiant", "Both", "555-0209", "nora.officiant@example.com", "Conduct ceremony", "Formal attire", ""),
    ("Emily Chambers", "Flower Girl", "Partner 1", "555-0210", "n/a", "Walk down aisle with petals", "Ivory dress", "Age 6"),
]

# -------------------------------------------------------- VENUE COMPARISON --
VENUES = [
    ("Silver Oak Estate", "Napa Valley, CA", 180, 12000, 9500, 6000, "Available", "Yes", "Yes", "Yes", "Yes", 5, "Selected venue — booked and deposit paid."),
    ("The Wilder Barn", "Sonoma, CA", 150, 9500, 8800, 4500, "Available", "Yes", "Yes", "Limited", "No", 4, "Rustic barn, strong runner-up."),
    ("Bellrose Gardens", "St. Helena, CA", 200, 14500, 11000, 7000, "Booked (Other Event)", "Yes", "Yes", "Yes", "Yes", 5, "Beautiful but unavailable on our date."),
    ("Ridgeview Terrace", "Calistoga, CA", 120, 8200, 7600, 3800, "Available", "No", "Yes", "Yes", "No", 3, "Ceremony space too small."),
    ("Orchard Hill Manor", "Yountville, CA", 160, 11000, 9000, 5500, "Available", "Yes", "Yes", "Yes", "Yes", 4, "Great option, slightly over budget."),
]

# ------------------------------------------------------------ FOOD & DRINKS --
FOOD_DRINKS = [
    ("Harvest Table Catering", "Herb-Crusted Goat Cheese Tartlet", "Appetizers", 150, 4.50, "Vegetarian", "Crowd favorite"),
    ("Harvest Table Catering", "Seared Scallop, Citrus Beurre Blanc", "Appetizers", 150, 6.00, "Contains Shellfish", ""),
    ("Harvest Table Catering", "Grilled Filet Mignon, Red Wine Jus", "Main Course", 90, 38.00, "Gluten-Free Option", "Most popular entrée"),
    ("Harvest Table Catering", "Pan-Seared Salmon, Lemon Dill", "Main Course", 45, 32.00, "Gluten-Free", ""),
    ("Harvest Table Catering", "Wild Mushroom Risotto", "Main Course", 30, 26.00, "Vegetarian, Vegan Option", ""),
    ("Harvest Table Catering", "Herb Roasted Chicken", "Main Course", 15, 28.00, "Gluten-Free", "Kids meal base"),
    ("Sugar & Sage Bakery", "Vanilla Bean Dessert Bites", "Dessert", 150, 3.00, "Contains Nuts", ""),
    ("Sugar & Sage Bakery", "Chocolate Salted Caramel Tart", "Dessert", 150, 3.50, "", ""),
    ("Sugar & Sage Bakery", "4-Tier Wedding Cake (Vanilla & Lemon)", "Cake", 150, 7.50, "", "Featured dessert table centerpiece"),
    ("Harvest Table Catering", "Signature Cocktail — Lavender Gin Fizz", "Drinks", 150, 12.00, "", "Named after the couple"),
    ("Harvest Table Catering", "Wine Service (Red & White)", "Drinks", 150, 18.00, "", "Local Napa selections"),
    ("Harvest Table Catering", "Non-Alcoholic Bar", "Drinks", 150, 6.00, "", ""),
    ("Harvest Table Catering", "Late Night Slider Station", "Late Night Food", 120, 8.00, "", "Served at 10pm"),
    ("Harvest Table Catering", "Late Night Popcorn Bar", "Late Night Food", 120, 3.00, "", ""),
]

# --------------------------------------------------------------- HONEYMOON --
HONEYMOON = [
    ("Flights", 1800, 1750, 1750, 0, "2027-06-01"),
    ("Hotel", 3200, 2900, 1500, 1400, "2027-06-20"),
    ("Transportation", 400, 380, 380, 0, "2027-06-05"),
    ("Food", 1200, 900, 300, 600, "2027-06-25"),
    ("Activities", 900, 320, 100, 220, "2027-06-25"),
    ("Shopping", 300, 200, 0, 200, "2027-06-25"),
    ("Travel Insurance", 150, 150, 150, 0, "2027-05-15"),
    ("Miscellaneous", 50, 0, 0, 0, "2027-06-25"),
]

# ---------------------------------------------------------------- CONTACTS --
CONTACTS = [
    ("Marisol Herrera", "Silver Oak Estate", "Venue", "Events Manager", "707-555-0142", "events@silveroakestate.com", "silveroakestate.com", "4200 Vineyard Rd, Napa, CA", ""),
    ("Grant Whitfield", "Harvest Table Catering", "Catering", "Owner", "707-555-0198", "grant@harvesttablecatering.com", "harvesttablecatering.com", "120 Main St, Napa, CA", ""),
    ("Ines Callahan", "Golden Hour Photo Co.", "Photography", "Lead Photographer", "415-555-0110", "ines@goldenhourphoto.co", "goldenhourphoto.co", "88 Market St, San Francisco, CA", ""),
    ("Theo Marsh", "Lumière Films", "Videography", "Director", "415-555-0177", "hello@lumierefilms.com", "lumierefilms.com", "220 Bay Ave, San Francisco, CA", ""),
    ("Camille Everly", "Maison Everly Bridal", "Attire", "Designer", "707-555-0133", "camille@maisoneverly.com", "maisoneverly.com", "15 Rose Ln, St. Helena, CA", ""),
    ("Rosalind Abernathy", "Wildroot Floral Design", "Flowers", "Owner", "707-555-0121", "rosalind@wildrootfloral.com", "wildrootfloral.com", "77 Garden Way, Napa, CA", ""),
    ("Devon Okafor", "Skyline Sound Co.", "Entertainment", "DJ / MC", "415-555-0155", "devon@skylinesound.com", "skylinesound.com", "310 Sound Dr, Oakland, CA", ""),
    ("Nora Fitzgerald (Officiant)", "N/A", "Officiant", "Celebrant", "707-555-0128", "nora.officiant@example.com", "", "Napa, CA", "Family friend, ordained online"),
    ("Willa Chambers", "Sugar & Sage Bakery", "Cake", "Head Baker", "707-555-0119", "willa@sugarandsage.com", "sugarandsage.com", "45 Baker St, Napa, CA", ""),
    ("Colin Sinclair", "Napa Ridge Inn", "Accommodation", "Reservations Manager", "707-555-0192", "colin@naparidgeinn.com", "naparidgeinn.com", "900 Ridge Rd, Napa, CA", ""),
]

# --------------------------------------------------------------- TIMELINE --
TIMELINE = [
    ("14:00", "Hair & Makeup Begins", "Bridal Suite, Silver Oak Estate", "Bloom Beauty Studio", "Bloom Beauty Studio", "Bride and bridesmaids"),
    ("15:30", "First Look & Portraits", "Vineyard Overlook", "Ines Callahan", "Golden Hour Photo Co.", ""),
    ("16:00", "Guests Arrive", "Ceremony Lawn", "Wedding Planner", "Silver Oak Estate", ""),
    ("16:30", "Ceremony", "Ceremony Lawn", "Nora Fitzgerald", "N/A", "20-minute ceremony"),
    ("17:00", "Cocktail Hour", "Terrace", "Harvest Table Catering", "Harvest Table Catering", "String quartet plays"),
    ("18:00", "Reception Entrance & First Dance", "Reception Pavilion", "Devon Okafor", "Skyline Sound Co.", ""),
    ("18:15", "Dinner Service", "Reception Pavilion", "Harvest Table Catering", "Harvest Table Catering", "Plated 3-course dinner"),
    ("19:30", "Toasts & Speeches", "Reception Pavilion", "William Foster", "N/A", "Best man & maid of honor"),
    ("20:00", "Cake Cutting", "Reception Pavilion", "Sugar & Sage Bakery", "Sugar & Sage Bakery", ""),
    ("20:15", "Open Dancing", "Reception Pavilion", "Devon Okafor", "Skyline Sound Co.", ""),
    ("22:00", "Late Night Snacks", "Reception Pavilion", "Harvest Table Catering", "Harvest Table Catering", "Slider & popcorn station"),
    ("23:00", "Send-Off", "Main Entrance", "Wedding Planner", "Valley Transit Co.", "Sparkler send-off"),
]

# ----------------------------------------------------------- GIFT TRACKER --
GIFT_TYPES = ["Physical Gift", "Cash Gift", "Gift Card", "Experience", "Registry Item"]
GIFTS_ITEMS = ["Set of Le Creuset Cookware", "Cash Gift", "$150 Amazon Gift Card", "Wine Country Tour Voucher",
               "Crystal Serving Bowl (Registry)", "Linen Bedding Set (Registry)", "Cash Gift", "KitchenAid Mixer (Registry)",
               "$100 Williams Sonoma Gift Card", "Hand-painted Ceramic Vase", "Cash Gift", "Weekend Spa Package"]


def build_gift_rows(n=26):
    names = full_names(n)
    rows = []
    for i, (f, l) in enumerate(names):
        gtype = GIFT_TYPES[i % len(GIFT_TYPES)]
        item = GIFTS_ITEMS[i % len(GIFTS_ITEMS)]
        value = [150, 200, 100, 250, 300][i % 5] if gtype != "Cash Gift" else [100, 150, 200, 250, 500][i % 5]
        thanked = "Yes" if i % 3 != 0 else "No"
        rows.append({
            "Guest Name": f"{f} {l}",
            "Gift": item,
            "Gift Type": gtype,
            "Value": value,
            "Thank You Sent": thanked,
            "Date Received": f"2027-0{(i % 6) + 1}-{(i % 27) + 1:02d}",
            "Notes": "",
        })
    return rows
