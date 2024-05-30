# Log generator for [OpenDominion Game](https://github.com/OpenDominion/OpenDominion)

Converted from the script included in sim file from [Yami-10/OD-Simulator](https://github.com/Yami-10/OD-Simulator)

# The idea
I work on Linux so don't have Excel and has no ability to generate `log.txt` from sim file.

This tool writen in [GO](https://go.dev/) can generate `log.txt` from Excel file
both local and downloaded from Google Sheet. It should work in any Operation System.

The code contains only final script generation logic. All other things left the same in Excel file.

Remember to open Excel sim before generation. ALl formulas should run and update their values with Excel.

# Usage
```
od_sim generate_log -sim OpenDominionSim.xlsm -result sim.txt
```

Get file from [Yami-10/OD-Simulator](https://github.com/Yami-10/OD-Simulator)

# Bug reports
If you see any issues or want an improvement, feel free to create an issue and describe the problem.

Write me in Discord @max_masterius
