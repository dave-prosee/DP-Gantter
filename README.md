# DP-Gantter
Excel add in to create and manage Gantt charts

Download the manifest file to sideload it (use home insert add in on a mac).
Once the add in is accepted the add in will be loaded from the github page, so a web connection is necessary to use it.

Some tips on usages:
- there is a button for new plannings. Still to be improved, but it will give you an idea how to use the add in
- for now, please use 100% zoom when (re-drawing)
- on the left, from column T, there are some columns that are either mandatory or for the users. This way you can create a data set that is filterable by your own data. The mandatory columns are ID, Name, From (calculated), Until (calculated), Duration (input), Do not start before (input), Dependency (input), Responsible, Progress (for a progress bar) and Color / Task type. 
- The Color/Type column decides over graph type and their colors. You can use the standard Excel fill color to color it. Tasks are the principal work to be done, a Phase will consolidate the duration of subordinate (indented) Tasks, and a Milestone will show up as a diamond.
- Dependency relate to the IDs of predecessors that need to finish first. You can add as many as you like, seperated by a comma and a space (to not make it a decimal number)
- IDs should be unique and should never be empty. They should also not be duplicated. It is checked for and the drawing halts on a misconfiguration
- Do not put anything below your planning, not even colored cells, as colored cells are recognized as used cells and then also put into the calculations.

