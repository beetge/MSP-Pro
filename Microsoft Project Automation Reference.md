# Microsoft Project Automation Reference

## Scope and platform boundaries

This report focuses on the **desktop Microsoft Project automation stack**: Project Standard, Project Professional, and the Project Online Desktop Client, because those are the Product editions covered by the Project VBA object model, macro recording support, and the desktop command model in Microsoft’s documentation. Microsoft also documents two broader extension paths: **VSTO add-ins** for local desktop deployment, and **Office task-pane add-ins** for Project on Windows. For task-pane add-ins, Microsoft explicitly says that Project uses the **Office JavaScript Common API** and **does not have a Project-specific JavaScript API**. citeturn35view0turn34view3turn34view4turn34view5

At the object-model level, Microsoft’s Project reference says the Project VBA documentation covers the objects, properties, methods, and events in the Project object model, and points developers to the official object maps for the Application/Projects, Tasks, Resources, and Calendars hierarchies. That makes the desktop VBA model the primary “feature-to-object” source of truth for deep schedule automation. citeturn38search5turn28view7turn17search1turn17search5turn28view2

## Feature-to-object mapping

The table below is a conservative, documentation-based map of the **main Microsoft Project feature areas** to their **primary VBA surfaces**.

| UI feature area | Primary VBA surface | Practical notes | Key source |
|---|---|---|---|
| Application shell, active window, selections, app options | `Application`, `ActiveProject`, `ActiveSelection`, `ActiveCell`, `Selection`, `Cell` | `Application` represents the entire Project app and exposes app-wide settings, top-level objects, and methods acting on views, selections, and editing. `Selection` and `Cell` give you the current UI context. | citeturn37view2turn24search0turn28view5turn24search10turn28view6 |
| Open project plan / document model | `Project`, `Projects` | `Project` represents one open project in the `Projects` collection and exposes project-level events such as `Open`, `BeforeSave`, `BeforeClose`, `Calculate`, and `Change`. | citeturn37view3turn38search8 |
| Tasks | `Project.Tasks`, `Task` | `Task` is the core scheduling entity; Microsoft documents task access through the `Tasks` collection and task-level properties such as `OutlineLevel`, `WBS`, constraints, baselines, and assignments. | citeturn29search10turn8search3turn30view7turn29search2 |
| Resources | `Project.Resources`, `Resource` | `Resource` is the master resource entity and carries calendars, assignments, custom fields, and baseline fields. | citeturn10search15turn23search1turn31view1 |
| Assignments and timephased work | `Task.Assignments`, `Resource.Assignments`, `Assignment`, `Assignment.TimeScaleData` | `Assignment` is the task-resource allocation object. It is the key surface for timephased automation, work contours, actuals, and baseline assignment data. | citeturn28view3turn23search1turn23search4turn28view4 |
| Dependencies and linking | `Task.LinkPredecessors`, `Application.LinkTasks`, `Application.UnlinkTasks` | Use `Task.LinkPredecessors` when you want object-level control over task relationships. Use `LinkTasks` / `UnlinkTasks` when you want the ribbon command equivalent against the current selection. | citeturn29search3turn26view0turn26view1 |
| Constraints and status date | `Task.ConstraintType`, `Task.ConstraintDate`, `Project.StatusDate` | Microsoft explicitly ties several constraint types to `ConstraintDate`. `StatusDate` is a project-level property. | citeturn29search2turn29search1turn29search4 |
| Baselines | `Application.BaselineSave`, `Application.BaselineClear`, baseline task/resource/assignment fields | Baseline automation is centered on save/clear methods and baseline fields on tasks, resources, and assignments. | citeturn30view4turn30view5turn10search2turn10search7turn23search8 |
| Outline and WBS | `Task.OutlineLevel`, `Application.OutlineIndent`, `Application.OutlineOutdent`, `Application.OutlineShowTasks`, `Task.WBS`, `Application.WBSCodeMaskEdit` | Project separates the outline hierarchy from WBS code-mask management. Both are scriptable. | citeturn8search3turn26view2turn26view3turn30view6turn30view7turn26view10 |
| Custom fields, project fields, enterprise fields | `Application.FieldNameToFieldConstant`, `Task.SetField`, `Resource.SetField`, `Project.ProjectSummaryTask` | Use `FieldNameToFieldConstant` to resolve local or enterprise field names to field IDs. Project-level custom fields are accessed through the project summary task. Enterprise project/resource fields use `SetField` / `GetField`. | citeturn26view14turn31view0turn31view1turn26view15 |
| Views | `Project.ViewsSingle`, `ViewSingle`, `Application.ViewApply`, `ViewsSingle.Add` | `ViewSingle` is the single-pane view object, with properties for table, filter, group, screen, and type. | citeturn26view7turn25view0turn27search0turn25view1turn27search13 |
| Filters | `Filter`, `Filters`, `Application.FilterApply` | Filters are first-class objects and can be applied directly by object or by application method. | citeturn27search2turn27search18turn26view4 |
| Groups | `Group`, `TaskGroups`, `ResourceGroups`, `Application.GroupApply` | Groups are first-class definitions, grouped by task vs resource collections. | citeturn25view2turn26view5 |
| Tables | `Table`, `Tables`, `Application.TableApply` | View tables are first-class objects, with `TableFields`, menu visibility, and table type. | citeturn25view3turn27search23turn26view6turn27search11 |
| Calendars | `Calendar`, `Calendars`, `Project.BaseCalendars` | Use `Project.BaseCalendars` to retrieve base calendars. A subtle but important distinction: `Application.BaseCalendars` is the **Change Working Time dialog command**, not the collection accessor. | citeturn28view0turn28view1turn31view4turn31view5 |
| Timeline | `Application.TaskOnTimeline`, `Application.TaskOnTimelineEx`, `Application.TimelineGotoSelectedTask`, `Application.RemoveTimelineBar` | Timeline automation is substantial. `TaskOnTimelineEx` and `RemoveTimelineBar` were introduced in Office 2016, and `TimelineGotoSelectedTask` is explicitly documented as the command equivalent for “Go to Selected Task.” | citeturn11search0turn26view12turn26view11turn26view13turn11search9 |
| Reports | `Project.Reports`, `Report`, `Reports.Add`, `Report.Shapes`, `ReportTable` | `Project.Reports` contains **custom** reports, not built-in ones. Reports can contain `Shape`, `Chart`, and `ReportTable` content. Microsoft also states that **macro recording for reports is not implemented**. | citeturn25view5turn25view4turn32search3turn32search2turn32search1 |
| Ribbon and legacy command bars | `Project.SetCustomUI`, `Application.CommandBars`, `Project.CommandBars` | Modern ribbon customization is XML-based through `SetCustomUI`. `CommandBars` remains available, but Microsoft notes that command bars were superseded in some Office apps by the Fluent ribbon. | citeturn25view6turn25view7turn15search4turn15search2 |
| Organizer and global template | `Application.OrganizerMoveItem`, `Global.mpt` | Use the Organizer to move views, calendars, reports, and similar definitions. `Global.mpt` is a global formatting/customization template and cannot store tasks, resources, or assignments. | citeturn25view8turn34view1turn34view2 |
| Import/export maps | `Application.MapEdit`, `Project.MapList`, `FileOpenEx(..., Map:=...)`, `FileSaveAs(..., Map:=...)` | Project supports named import/export maps; Microsoft recommends the `Map` argument over older compatibility arguments such as `Table`, `Sheet`, and `TaskInformation` for non-Project file formats. | citeturn17search14turn17search11turn17search4turn30view2 |

## Ribbon command and shortcut mapping

This table maps **common ribbon/menu commands** to **direct or closest VBA calls** that Microsoft documents.

| UI command | Direct or closest VBA call | Notes | Key source |
|---|---|---|---|
| File > New | `Application.FileNew` | Creates a new project; can optionally use a template. | citeturn30view0 |
| File > Open / Import | `Application.FileOpenEx` | Opens a project or imports data. Supports multiple formats and Project Server/Draft scenarios. | citeturn30view1 |
| File > Save As / Export | `Application.FileSaveAs` | Saves under a new name or exports data to another file format; supports map-based export. | citeturn30view2 |
| File > Properties | `Application.FileProperties` plus `BuiltInDocumentProperties` / `CustomDocumentProperties` | `FileProperties` opens the dialog; document-property collections are the programmable route. | citeturn14search0turn30view3turn31view2turn31view3 |
| File > Organizer | `Application.OrganizerMoveItem` | Programmatic Organizer equivalent. | citeturn25view8 |
| Task > Link Selected Tasks | `Application.LinkTasks` | Selection-based ribbon equivalent. | citeturn26view0 |
| Task > Unlink Tasks | `Application.UnlinkTasks` | Selection-based ribbon equivalent. | citeturn26view1 |
| Task > Indent | `Application.OutlineIndent` | Indents the selected task(s). | citeturn26view2 |
| Task > Outdent | `Application.OutlineOutdent` | Promotes the selected task(s). | citeturn26view3 |
| View > Filter | `Application.FilterApply` or `Filter.Apply` | Application method or first-class filter object. | citeturn26view4turn27search18 |
| View > Group By | `Application.GroupApply` | Applies the named group. | citeturn26view5 |
| View > Table | `Application.TableApply` or `Table.Apply` | Applies the named table to the active view. | citeturn26view6turn25view3 |
| View > Change View | `Application.ViewApply` or `ViewSingle.Apply` | `ViewApply` is the general command; `ViewSingle.Apply` is object-based. | citeturn27search0turn27search13 |
| Resource > Level All | `Application.LevelNow True` | Levels overallocated resources. | citeturn26view8 |
| Resource > Leveling Options | `Application.LevelingOptionsEx ...` | Scripts the leveling options instead of opening the dialog. | citeturn26view9turn16search8 |
| Project > Change Working Time | `Application.BaseCalendars` | This is the dialog-command equivalent for Change Working Time. For calendar retrieval/manipulation, use `Project.BaseCalendars`. | citeturn31view5turn31view4 |
| Timeline > Add or remove task from timeline | `Application.TaskOnTimeline` / `Application.TaskOnTimelineEx` | `TaskOnTimelineEx` adds support for custom timelines and bar indexes. | citeturn11search0turn26view12 |
| Timeline option menu > Go to Selected Task | `Application.TimelineGotoSelectedTask` | Microsoft explicitly says this corresponds to that option-menu command. | citeturn26view11 |
| Timeline > Remove bar | `Application.RemoveTimelineBar` | Timeline bar management. | citeturn26view13 |
| Project / Format > WBS definition | `Application.WBSCodeMaskEdit` | Edits the WBS code mask. | citeturn26view10 |
| Report Tools > Design > Add custom report | `ActiveProject.Reports.Add` | Creates a custom report and switches to Report Tools / Design. | citeturn25view4turn32search3 |
| Report Tools > Design > Table | `Application.Table` | Important naming collision: this creates a **report table**, not a view table. | citeturn32search8turn32search1 |

The shortcut map below is intentionally strict: it includes only shortcuts Microsoft explicitly documents for Project, plus the closest documented VBA call when there is one.

| Shortcut | UI effect | Documented VBA mapping | Notes | Key source |
|---|---|---|---|---|
| `Ctrl+C` | Copy selection | `Application.EditCopy` | Direct command equivalent. | citeturn2view8turn39search4 |
| `Ctrl+X` | Cut selection | `Application.EditCut` | Direct command equivalent. | citeturn22view5turn39search8 |
| `Ctrl+V` | Paste into active selection | `Application.EditPaste` | Direct command equivalent. | citeturn22view3turn39search5 |
| `Delete` or `Ctrl+Delete` | Delete selected data or row | `Application.EditDelete` | Exact behavior depends on whether a row, column, or active-cell row is selected. | citeturn22view5turn39search3 |
| `Ctrl+F2` | Link tasks | `Application.LinkTasks` | Direct ribbon/menu equivalent. | citeturn22view3turn26view0 |
| `Ctrl+Shift+F2` | Unlink tasks | `Application.UnlinkTasks` | Direct ribbon/menu equivalent. | citeturn22view4turn26view1 |
| `Alt+Shift+Right Arrow` | Indent selected task | `Application.OutlineIndent` | Direct outline action equivalent. | citeturn22view5turn26view2 |
| `Alt+Shift+Left Arrow` | Outdent selected task | `Application.OutlineOutdent` | Direct outline action equivalent. | citeturn22view6turn26view3 |
| `Shift+F2` | Open task information in supported timeline contexts | No one-line VBA command verified in the reviewed Project docs | For automation, use object properties directly instead of driving the dialog. | citeturn22view5turn29search10 |
| `F2` | Activate edit mode for a cell | No one-line VBA command verified in the reviewed Project docs | This is primarily a cell-editing UI command. | citeturn22view5turn28view6 |

A practical consequence of this mapping is that **Project exposes two different automation styles**. Some commands are **selection-driven UI verbs** such as `LinkTasks`, `OutlineIndent`, and `EditPaste`. Others are **object-driven model operations** such as `Task.LinkPredecessors`, `Task.SetField`, or `Assignment.TimeScaleData`. In real automation, object-driven code is usually more robust, while selection-driven code is more useful when you intentionally want to mimic ribbon behavior. citeturn26view0turn26view2turn39search5turn29search3turn31view0turn28view4

## Macro procedures and high-value code patterns

Microsoft’s own Project support article confirms that Project has a **macro recorder**. To record a macro, go to **View > Macros > Record Macro**, give the macro a name, optionally assign a shortcut key, choose whether to store it in **Global File** or **This Project**, choose relative or absolute row/column recording behavior, perform the actions, and then stop recording. Microsoft also warns that macros can contain viruses and points users to macro security and trusted-publisher precautions. citeturn35view0

A useful way to think about Project macros is this: use the recorder to discover **UI-driven commands**, but use the object model when you need **repeatable, parameterized, testable** automation. Microsoft’s own docs point to `Task`, `Resource`, `Assignment`, `FieldNameToFieldConstant`, `SetField`, `BaselineSave`, `LevelingOptionsEx`, and timeline/report objects for that deeper automation layer. citeturn29search10turn10search15turn28view3turn26view14turn31view0turn31view1turn30view4turn26view9turn25view4

The first pattern below is the most reliable “schedule build” macro shape: create or open a project, add tasks through the object model, and only use selection-based methods for actions that are inherently UI-style operations.

```vb
Option Explicit

Public Sub BuildSimplePlan()
    Dim pj As Project
    Dim t1 As Task
    Dim t2 As Task
    Dim t3 As Task

    On Error GoTo CleanFail

    Application.ScreenUpdating = False

    If Projects.Count = 0 Then
        Application.FileNew
    End If

    Set pj = ActiveProject

    Set t1 = pj.Tasks.Add("Design")
    t1.Duration = "5d"

    Set t2 = pj.Tasks.Add("Build")
    t2.Duration = "10d"

    Set t3 = pj.Tasks.Add("Test")
    t3.Duration = "4d"

    ' Object-model dependency creation is usually more robust than selection-driven linking.
    t2.LinkPredecessors t1
    t3.LinkPredecessors t2

    ' Save the initial baseline for all tasks.
    Application.BaselineSave All:=True

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "BuildSimplePlan failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

This pattern is grounded in Microsoft’s documentation for `Project`, `Tasks.Add`, `Task.LinkPredecessors`, and `Application.BaselineSave`. citeturn37view3turn29search10turn29search3turn30view4

The next pattern is the standard **custom-field resolution** approach for local and enterprise fields: resolve the field ID with `FieldNameToFieldConstant`, then read or write the field through `SetField` / `GetField`, using `ProjectSummaryTask` for project-level custom fields when needed. citeturn26view14turn31view0turn31view1turn26view15

```vb
Option Explicit

Public Sub SetAProjectCustomField()
    Dim fieldId As Long
    Dim pst As Task

    On Error GoTo CleanFail

    ' Example: local or enterprise project-level custom field.
    fieldId = Application.FieldNameToFieldConstant("Sponsor", pjProject)
    Set pst = ActiveProject.ProjectSummaryTask

    pst.SetField FieldID:=fieldId, Value:="PMO"

    MsgBox "Project field updated.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "SetAProjectCustomField failed: " & Err.Description, vbExclamation
End Sub
```

For **resource leveling**, Microsoft documents a clean two-step approach: first set options with `LevelingOptionsEx`, then call `LevelNow`. That is dramatically more maintainable than trying to drive the Leveling Options dialog manually. citeturn26view9turn26view8

```vb
Option Explicit

Public Sub LevelWholeProject()
    On Error GoTo CleanFail

    Application.LevelingOptionsEx _
        Automatic:=False, _
        DelayInSlack:=False, _
        AutoClearLeveling:=True, _
        Order:=pjLevelImmediate, _
        LevelEntireProject:=True, _
        LevelIndividualAssignments:=True, _
        LevelingCanSplit:=True

    Application.LevelNow True
    MsgBox "Leveling complete.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "LevelWholeProject failed: " & Err.Description, vbExclamation
End Sub
```

For **timeline automation**, the methods worth knowing are `TaskOnTimeline`, `TaskOnTimelineEx`, `TimelineGotoSelectedTask`, and `RemoveTimelineBar`. Microsoft explicitly notes that `TaskOnTimelineEx` and `RemoveTimelineBar` were introduced in Office 2016, and that `TimelineGotoSelectedTask` errors if the active view is not Timeline or if a single timeline task is not selected. citeturn26view12turn26view13turn26view11

```vb
Option Explicit

Public Sub AddTaskToTimeline()
    On Error GoTo CleanFail

    ' Adds task ID 1 to the default timeline.
    Application.TaskOnTimeline TaskID:=1, Remove:=False

    MsgBox "Task added to timeline.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "AddTaskToTimeline failed: " & Err.Description, vbExclamation
End Sub
```

## Event model, templates, and deployment surfaces

Project supports both **project-level events** and **application-level events**. Microsoft says project-level event procedures live on the `Project` object for any open project, while **application-level events require a class module with `WithEvents`**. Microsoft’s event guidance also says that if you automate Project from another application, you should register application-level event handlers **after** setting `Application.Visible = True`; otherwise, child objects like `ActiveProject` may not behave correctly. citeturn38search1turn37view2

Microsoft also documents a subtle but important **Global.mpt interaction rule**: event code in the global file can run unexpectedly or block project event behavior. If both project and global event code exist, project code generally wins for that project event. If the project does not contain that event code but `Global.mpt` does, the global event runs. citeturn38search1

`Global.mpt` itself is best understood as a **global formatting and customization template**, not a schedule template. Microsoft’s file-format documentation says the global file contains formatting information for all projects and **cannot store tasks, resources, or assignments**. The support documentation for the Organizer adds that you can use it to copy reports, calendars, views, groups, filters, and similar elements into the Global template so they become available across projects. Microsoft also notes that Project can automatically add new views, tables, filters, and groups to the global template, depending on the Advanced option setting. citeturn34view2turn34view1

For deployment, Microsoft currently documents three main technical surfaces. **VBA macros** and **VSTO add-ins** are for local desktop use with Project Standard and Project Professional. **Office task-pane add-ins** can be distributed through the Office add-in model for Project on Windows, but Microsoft explicitly says that Project add-ins use the **Office JavaScript Common API** and do **not** have a Project-specific JavaScript API. In practice, that means VBA/COM/VSTO remains the deeper path for schedule-centric automation, while Office add-ins are better for pane UI, cross-service integration, and lightweight app experiences. citeturn34view3turn34view4turn34view5

On the security side, Project exposes the `Application.AutomationSecurity` property, which controls how macros behave when files are opened **programmatically**. Microsoft documents three important modes: `msoAutomationSecurityByUI`, `msoAutomationSecurityForceDisable`, and `msoAutomationSecurityLow`, with the default being the Trust Center behavior. Project also exposes `Project.VBProject`, but Microsoft says using that surface requires a reference to the **VBIDE** library. citeturn37view0turn37view1

## Gaps, caveats, and practical recommendations

The biggest caveat in the official documentation is that **not every UI feature has a clean one-to-one VBA method**. This is especially true for some dialog navigation shortcuts and some newer report-design interactions. Microsoft explicitly states that while Project has a macro recorder, **report add/edit actions are not recorded**. That is why report automation often starts with `Reports.Add` and then shifts into `Shape`, `Chart`, and `ReportTable` automation instead of relying on recorded code. citeturn35view0turn25view4turn32search3turn32search2turn32search1

A second caveat is that the same word can mean different things on different surfaces. For example, **table automation** in normal task/resource views uses the `Table` / `Tables` object family, while **report tables** are `ReportTable` shapes and can also be created with `Application.Table` when a report is active. Likewise, `Project.BaseCalendars` is the calendar collection accessor, while `Application.BaseCalendars` is the dialog-style command equivalent for Change Working Time. citeturn25view3turn32search8turn32search1turn31view4turn31view5

A third caveat is that **selection-based automation is context-sensitive**. `LinkTasks` depends on task selection in supported task views. `TimelineGotoSelectedTask` throws run-time error 1100 if the Timeline view is not active or if exactly one timeline task is not selected. This is one reason enterprise-grade Project macros tend to favor object-model calls such as `Task.LinkPredecessors`, `SetField`, and `TimeScaleData` over UI-style verbs whenever the object model exposes the same business operation. citeturn26view0turn26view11turn29search3turn31view0turn28view4

The strongest practical recommendations from the evidence are straightforward. Prefer **object-driven code** for schedule logic. Use the macro recorder mainly to discover **UI verbs** and their parameter order. Store cross-project macros in **Global File** only when you truly want them global. Use **OrganizerMoveItem** and the Organizer workflow for reusable views/tables/reports/calendars. Use **`FieldNameToFieldConstant` + `SetField`** for custom-field automation. Use **`LevelingOptionsEx` + `LevelNow`** for leveling. Use **`TaskOnTimelineEx`** for richer timeline control. And if you need app UI or cloud-service integration more than deep schedule control, use a **Project task-pane add-in**—but plan around the fact that it is a **Common API** add-in, not a full schedule object model. citeturn35view0turn25view8turn26view14turn31view0turn31view1turn26view9turn26view8turn26view12turn34view4turn34view5

Open questions and limitations remain. I did **not** verify one-to-one VBA mappings for every keyboard shortcut or every backstage/ribbon navigation key-tip combination, because Microsoft’s Project documentation is uneven there. I also excluded undocumented “folk mappings” that are common in blog posts but not confirmed in Microsoft’s own Project documentation. The tables above are therefore intentionally **conservative rather than exhaustive**, which is the safest way to build a reliable automation reference from primary Microsoft sources. citeturn2view8turn38search5