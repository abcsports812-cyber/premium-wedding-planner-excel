import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties

from lists_settings import build_lists_sheet, build_settings_sheet
from sheets import (
    build_budget, build_vendors, build_guest_list, build_payments, build_checklist,
    build_wedding_party, build_venue_comparison, build_food_drinks, build_honeymoon,
    build_contacts, build_gift_tracker, build_timeline,
)
from special_sheets import build_seating, build_calendar, build_start_here
from dashboard import build_dashboard


def build_workbook(demo: bool):
    wb = Workbook()
    wb.remove(wb.active)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)

    ranges = build_lists_sheet(wb)
    build_start_here(wb, demo)

    budget_ws, b0, b1 = build_budget(wb, ranges, demo)
    vendors_ws, v0, v1 = build_vendors(wb, ranges, demo)
    guest_ws, g0, g1 = build_guest_list(wb, ranges, demo)
    build_seating(wb, g0, g1, demo)
    build_timeline(wb, ranges, demo)
    checklist_ws, c0, c1 = build_checklist(wb, ranges, demo)
    payments_ws, p0, p1 = build_payments(wb, ranges, demo)
    build_wedding_party(wb, ranges, demo)
    build_venue_comparison(wb, ranges, demo)
    build_food_drinks(wb, ranges, demo)
    honeymoon_ws, h0, h1 = build_honeymoon(wb, ranges, demo)
    build_contacts(wb, ranges, demo)
    build_gift_tracker(wb, ranges, demo)
    build_calendar(wb, c0, c1, p0, p1, demo)
    build_settings_sheet(wb, demo)

    refs = {
        "Budget": (b0, b1), "Vendors": (v0, v1), "Guest List": (g0, g1),
        "Payments": (p0, p1), "Master Checklist": (c0, c1), "Honeymoon Budget": (h0, h1),
    }
    build_dashboard(wb, refs, demo)

    order = ["START HERE", "Dashboard", "Budget", "Vendors", "Guest List", "Seating Plan",
             "Timeline", "Master Checklist", "Payments", "Wedding Party", "Venue Comparison",
             "Food & Drinks", "Honeymoon Budget", "Contacts", "Gift Tracker", "Calendar",
             "Settings", "Lists"]
    wb._sheets = [wb[name] for name in order]
    wb.active = 0
    return wb


if __name__ == "__main__":
    out_dir = os.path.join(os.path.dirname(__file__), "..", "dist")
    os.makedirs(out_dir, exist_ok=True)

    demo_wb = build_workbook(True)
    demo_path = os.path.join(out_dir, "Premium_Wedding_Planner_DEMO.xlsx")
    demo_wb.save(demo_path)
    print("Saved", demo_path)

    blank_wb = build_workbook(False)
    blank_path = os.path.join(out_dir, "Premium_Wedding_Planner_BLANK.xlsx")
    blank_wb.save(blank_path)
    print("Saved", blank_path)
