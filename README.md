# Ford After-Sales Service Improvement 2570

This repository contains a PowerPoint generator for a 2-slide executive presentation.

## Files
- `create_presentation.py` — creates the PowerPoint file

## To generate the presentation
1. Install Python 3.8+
2. Install the PPTX dependency:
   ```bash
   pip install python-pptx
   ```
3. Run the script:
   ```bash
   python create_presentation.py
   ```
4. The output file will be created as:
   `Ford_After_Sales_2570_2_Slides.pptx`

## Font behavior
- The script tries to use `Ford F-1 Light`.
- If that font is not available on the machine, Microsoft PowerPoint and the Python library will fall back to a nearby font like Arial.

## Notes
- This is a 2-slide presentation oriented toward executive review and approval.
- Content covers: key insights, priority initiatives, quick wins, timeline, and KPI targets.
