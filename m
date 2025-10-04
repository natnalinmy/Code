ClearCollect(
    colLine,
    MASTER_LINE
);
Set(
    _dateSelected,
    Today()
);

Set(
    varTaskTitle,
    Param("task_title")
);


Set(
    _firstDayOfDaySelected,
    Today()
);
Set(
    _firstDayOfMonth,
    DateAdd(
        Today(),
        1 - Day(Today()),
        TimeUnit.Days
    )
);
Set(
    _firstDayInView,
    DateAdd(
        _firstDayOfMonth,
        -(Weekday(_firstDayOfMonth) - 2 + 1),
        TimeUnit.Days
    )
);
Set(
    _lastDayOfMonth,
    DateAdd(
        DateAdd(
            _firstDayOfMonth,
            1,
            TimeUnit.Months
        ),
        -1,
        TimeUnit.Days
    )
);
Set(
    _lastDayOfMonth,
    DateAdd(
        _lastDayOfMonth,
        86399,
        TimeUnit.Seconds
    )
);
// Adds 23 hours, 59 minutes, and 59 seconds

Set(
    _minDate,
    DateAdd(
        _firstDayOfMonth,
        -(Weekday(_firstDayOfMonth) - 2 + 1),
        TimeUnit.Days
    )
);
Set(
    _maxDate,
    DateAdd(
        DateAdd(
            _firstDayOfMonth,
            -(Weekday(_firstDayOfMonth) - 2 + 1),
            TimeUnit.Days
        ),
        40,
        TimeUnit.Days
    )
);
// Set the current date to the first day of the week 
Set(
    _firstDayOfWeek,
    DateAdd(
        Today(),
        1 - Weekday(
            Today(),
            StartOfWeek.Sunday
        ),
        TimeUnit.Days
    )
);
// Set the last day of the week 
Set(
    _lastDayOfWeek,
    DateAdd(
        _firstDayOfWeek,
        6,
        TimeUnit.Days
    )
);
// Set the minimum date (start of the week) and maximum date (end of the week) for display purposes 
Set(
    _minDateWeek,
    _firstDayOfWeek
);
Set(
    _maxDateWeek,
    _lastDayOfWeek
);
Set(
    varUser,
    First(
        Filter(
            MASTER_EMPLOYEE,
            emp_email = User().Email
            // emp_email = "yanisa.taveemunputtakarn.a4k@ap.denso.com"

        )
    )
);
//Set df calendar mode
Set(
    Month,
    true
) And Set (
    Weekday,
    false
) And Set(
    Day,
    false
);
//Set df record edit
Set(
    lastSubmittedID,
    0
);



Notify(
    "กำลังรอโหลดข้อมูล....",
    NotificationType.Information
);
// ----------- Permission TA Collection ----------------
Clear(clPMS_TA_0);
ForAll(
    MASTER_LINE,
    ForAll(
        ThisRecord.rep_ta,
        Collect(
            clPMS_TA_0,
            {
                mail: Email,
                PD: pd,
                Line: line_key
            }
        )
    )
);
// ----------- Permission TL Collection ----------------
Clear(clPMS_TL_0);
ForAll(
    MASTER_LINE,
    ForAll(
        ThisRecord.rep_tl,
        Collect(
            clPMS_TL_0,
            {
                mail: Email,
                PD: pd,
                Line: line_key
            }
        )
    )
);
// ----------- Permission SV Collection ----------------
Clear(clPMS_SV_0);
ForAll(
    MASTER_LINE,
    Collect(
        clPMS_SV_0,
        {
            mail: ThisRecord.rep_sv.Email,
            PD: pd,
            Line: line_key
        }
    )
);
// ----------- Permission AM Collection ----------------
Clear(clPMS_AM_0);
ForAll(
    MASTER_LINE,
    Collect(
        clPMS_AM_0,
        {
            mail: ThisRecord.rep_am.Email,
            PD: pd,
            Line: line_key
        }
    )
);
// ----------- Permission MGR Collection ----------------
Clear(clPMS_MGR_0);
ForAll(
    MASTER_LINE,
    Collect(
        clPMS_MGR_0,
        {
            mail: ThisRecord.rep_mgr.Email,
            PD: pd,
            Line: line_key
        }
    )
);
// ----------- Permission GM Collection ----------------
Clear(clPMS_GM_0);
ForAll(
    MASTER_LINE,
    Collect(
        clPMS_GM_0,
        {
            mail: ThisRecord.rep_gm.Email,
            PD: pd,
            Line: line_key
        }
    )
);
//------------------------------------------------
ClearCollect(
    colPermission_0,
    clPMS_GM_0,
    clPMS_MGR_0,
    clPMS_AM_0,
    clPMS_SV_0,
    clPMS_TL_0,
    clPMS_TA_0
);
//----------------Collection All---------------
Collect(
    clPMS_All,
    Distinct(
        Filter(
            colPermission_0,
            mail = User().Email
        ),
        PD
    )
);
If(
    IsBlank(varFilterPD),
    Set(
        varFilterPD,
        First(
            Distinct(
                clPMS_All,
                Value
            )
        ).Value
    ),
    varFilterPD
);
If(
    IsBlank(varSectionCode),
    Set(
        varSectionCode,
        LookUp(
            colLine,
            line_key = First(
                Filter(
                    colPermission_0,
                    mail = User().Email
                )
            ).Line
        ).section_code
    )
);
ClearCollect(
    col5m1e,
    Filter(
        '5M1E_Improvement',
        time_start >= _firstDayOfMonth && time_start <= _lastDayOfMonth && pd = varFilterPD && section_code = varSectionCode
    )
);
Collect (
    col5m1e,
    Sort(
        col5m1e,
        time_start
    ),
    SortOrder.Descending
);


ClearCollect(
    colEmp,
    Filter(
        MASTER_EMPLOYEE,
        position_tier = "T8" || position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1"
    )
);

ClearCollect(
    colEmpQA,
    Filter(
        MASTER_EMPLOYEE,
        Short_Dep = "QA",
        (position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1" || position_tier = "T1/C" || position_tier = "G1" || position_tier = "G1/A1" || position_tier = "G2" || position_tier = "G2/A2")
    )
);
ClearCollect(
    colEmpPE,
    Filter(
        MASTER_EMPLOYEE,
        Short_Dep = "PE",
        (position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1" || position_tier = "T1/C" || position_tier = "G1" || position_tier = "G1/A1" || position_tier = "G2" || position_tier = "G2/A2")
    )
);
ClearCollect(
    colEmpPD,
    Filter(
        MASTER_EMPLOYEE,
        Short_Dep = "PD",
        (position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1" || position_tier = "T1/C" || position_tier = "G1" || position_tier = "G1/A1" || position_tier = "G2" || position_tier = "G2/A2")
    )
);
ClearCollect(
    colEmpPC,
    Filter(
        MASTER_EMPLOYEE,
        Short_Dep = "PC",
        (position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1" || position_tier = "T1/C" || position_tier = "G1" || position_tier = "G1/A1" || position_tier = "G2" || position_tier = "G2/A2")
    )
);
ClearCollect(
    colEmpMT,
    Filter(
        MASTER_EMPLOYEE,
        Short_Dep = "MT",
        (position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1" || position_tier = "T1/C" || position_tier = "G1" || position_tier = "G1/A1" || position_tier = "G2" || position_tier = "G2/A2")
    )
);
ClearCollect(
    colEmpQC,
    Filter(
        MASTER_EMPLOYEE,
        Short_Dep = "QC",
        (position_tier = "T7" || position_tier = "T6" || position_tier = "T5" || position_tier = "T4" || position_tier = "T3" || position_tier = "T2" || position_tier = "T1" || position_tier = "T1/C" || position_tier = "G1" || position_tier = "G1/A1" || position_tier = "G2" || position_tier = "G2/A2")
    )
);

ClearCollect(
    DropdownlinenameItems,
    {
        line_name: "All",
        pd: varFilterPD,
        section_code: varSectionCode
    }
);
Collect(
    DropdownlinenameItems,
    Filter(
        MASTER_LINE,
        pd = varFilterPD && section_code = varSectionCode
    )
);
ClearCollect(
    colContent,
    '5M1E_Improvement_Content'
);

ClearCollect(
    colCodeRecordChoices,
    { DisplayValue: "All", DataValue: "All"},
    { DisplayValue: "M1 : Man", DataValue: "M1"},
    { DisplayValue: "M2 : Machine", DataValue: "M2"},
    { DisplayValue: "M3 : Materail", DataValue: "M3"},
    { DisplayValue: "M4 : Method", DataValue: "M4"},
    { DisplayValue: "M5 : Measurement", DataValue: "M5"},
    { DisplayValue: "E : Enviroment", DataValue: "E"}
);

// Filter variables for approve notification screen
Set(varFilterPD_Approve, "");
Set(varFilterSectionCode_Approve, "");  
Set(varFilterStartDate_Approve, _firstDayOfMonth);
Set(varFilterEndDate_Approve,Today());
Set(varFilterStatus_Approve, "All");
Set(varFilterCategory_Approve, "All");
Set(varShowFilters_Approve, false);

ClearCollect(
    CategoryOptions,
    {Display: "All", Value: "All"},
    {Display: "M1 : Man", Value: "M1"},
    {Display: "M2 : Machine", Value: "M2"},
    {Display: "M3 : Method", Value: "M3"},
    {Display: "M4 : Material", Value: "M4"},
    {Display: "M5 : Measurement", Value: "M5"},
    {Display: "E : Environment", Value: "E"}
);

Notify(
    "โหลดข้อมูลเสร็จสิ้น"
);

// Second Code //

