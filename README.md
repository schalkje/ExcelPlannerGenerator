# Excel Planner Generator

The goal is to generate an excel planning sheet that is meant for planning off-days in a scrum team.

## Absence
code|meaning
---|---
S|Sickness
V|Vacation
X|Working but not contributing to the team effort
0|Not working

## parameters
### Team
A list of teammembers and their working times in a week.

- **Name**: Name of the team member
- **Start date**: Start date of the team involvement of this team member
- **End date**: End date of the team involvement of this team member
- **Monday..Friday**: number of hours this resource is working in a day

### Sprint
Sprints are labeled yyyy-ss, where ss is the 0 based sprint number (01,02,..,24,25,..)

- **Sprint flip day**: day of the week for the sprint flip
- **Sprint duration**: duratin of a sprint in weeks
- **First sprint**: start day of the first sprint of the year

# Installation and usage

Python needs to be installed ([download](https://www.python.org/downloads/))
or using chocolatey:
``` 
choco install python
```

The prerequisite python libraries can be installed by:
``` 
pip install -r requirements.txt
```

