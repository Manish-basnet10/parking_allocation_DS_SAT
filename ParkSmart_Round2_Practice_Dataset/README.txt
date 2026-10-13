ParkSmart City Services — Round 2 practice dataset
=====================================================

IMPORTANT: SYNTHETIC PRACTICE DATA. This is generated for learning/testing and is NOT the official assessment dataset or real parking data.

Files:
- parking_sessions.csv: session_id, zone_id, zone, entry, exit, duration (minutes), fee, payment_id
- zones.csv: zone_id, location, capacity, zone_type
- payments.csv: payment_id, session_id, payment_mode, amount
- calendar.csv: date, weekday, holiday, event
- weather.csv: date, rain, temperature

Join keys:
- parking_sessions.zone_id -> zones.zone_id
- parking_sessions.payment_id -> payments.payment_id (or session_id)
- date derived from parking_sessions.entry -> calendar.date and weather.date

Deliberate quality issues:
- Missing exit timestamps
- Duplicate session rows
- Inconsistent zone labels
- Negative time intervals
- Capacity mismatch to investigate
- Missing duration values
- Duplicate payment rows

Notes:
- Inspect and document data quality before cleaning.
- Do not assume missing exits are invalid; they may be open sessions.
- Estimate occupancy using overlapping entry/exit intervals; session counts alone are not occupancy.
