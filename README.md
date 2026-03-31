# CampusFlow Analytics

Streamlit app that treats a multi-sheet Excel workbook as a simple relational database to analyze campus space efficiency and detect **Ghost Rooms** (scheduled rooms with low attendance vs capacity).

## Expected Excel sheets

Your `.xlsx` must include these sheets and columns:

- `Resources`: `Room_ID`, `Building`, `Type`, `Capacity`
- `Schedule`: `Slot_ID`, `Day`, `Time`, `Room_ID`, `Course_ID`
- `Utilization`: `Slot_ID`, `Actual_Attendance`, `Date`

## Run locally

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Notes

- The app will warn on **over-capacity** cases where `Actual_Attendance > Capacity`.
- “Efficiency Score” is computed per scheduled slot record as \(100 \times \frac{\text{Actual Attendance}}{\text{Capacity}}\), capped to 100 for display, with a separate over-capacity warning.

## Future Scope

- AI-based demand prediction
- Mobile dashboard
- Integration with real-time sensors
