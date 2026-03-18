'==============================================================================
' Filename:    modTypes.vba
' Purpose:     User Defined Types shared across all logistics modules
' Dependencies: None
' By:          T. Sciple, 03/17/2026
'==============================================================================

'##############################################################################
'  MODULE: modTypes  (Standard Module)
'  User Defined Types shared across all modules
'##############################################################################
Option Explicit

' ── Transport mode enumeration ──
Public Enum transportMode
    tmRail = 1
    tmTruck = 2
    tmBarge = 3
End Enum

' ── Raw material tank ──
Public Type RawTank
    tank_name       As String
    material_name   As String
    capacity_bbl    As Double
    inventory_bbl   As Double
    min_inv_bbl     As Double
End Type

' ── Blend tank ──
Public Type BlendTank
    tank_name       As String
    capacity_bbl    As Double
    inventory_bbl   As Double
End Type

' ── Blend recipe line ──
Public Type BlendRecipeLine
    blend_tank_name As String
    material_name   As String
    fraction        As Double
End Type

' ── Processing unit ──
Public Type ProcessingUnit
    unit_name       As String
    capacity_bbl_day As Double
    feed_source     As String       ' name of raw tank or blend tank
    product_name    As String
End Type

' ── Product tank ──
Public Type ProductTank
    tank_name       As String
    product_name    As String
    capacity_bbl    As Double
    inventory_bbl   As Double
    min_inv_bbl     As Double
End Type

' ── Unload spot config ──
Public Type UnloadSpotConfig
    mode_type       As transportMode
    mode_name       As String
    num_spots       As Long
    avg_unload_hrs  As Double
    bbl_per_load    As Double
End Type

' ── Load spot config ──
Public Type LoadSpotConfig
    mode_type       As transportMode
    mode_name       As String
    num_spots       As Long
    avg_load_hrs    As Double
    bbl_per_load    As Double
End Type

' ── Scheduled arrival ──
Public Type ScheduledArrival
    arrival_day     As Long
    mode_name       As String
    quantity_bbl    As Double
    material_name   As String
End Type

' ── Scheduled shipment ──
Public Type ScheduledShipment
    ship_day        As Long
    product_name    As String
    quantity_bbl    As Double
    mode_name       As String
End Type

' ── A single load being unloaded (tracks remaining time) ──
Public Type ActiveUnload
    material_name   As String
    bbl_remaining   As Double
    hours_remaining As Double
    spot_index      As Long
End Type

' ── A single load being loaded (tracks remaining time) ──
Public Type ActiveLoad
    product_name    As String
    bbl_remaining   As Double
    hours_remaining As Double
    spot_index      As Long
End Type

' ── Snapshot of state at one time step (for results) ──
Public Type StepSnapshot
    sim_step        As Long
    date_time       As Date
    raw_inventories() As Double     ' indexed same as raw_tanks array
    blend_inventories() As Double
    product_inventories() As Double
    unit_throughputs() As Double
    unloading_active As Boolean
    loading_active  As Boolean
    flag_text       As String
End Type

' ── Master simulation state ──
Public Type SimState
    ' Configuration
    run_duration_days   As Long
    time_step_hrs       As Double
    unload_on_weekends  As Boolean
    load_on_weekends    As Boolean
    start_date          As Date
    total_steps         As Long

    ' Infrastructure arrays
    raw_tanks()         As RawTank
    blend_tanks()       As BlendTank
    blend_recipes()     As BlendRecipeLine
    units()             As ProcessingUnit
    product_tanks()     As ProductTank
    unload_spots()      As UnloadSpotConfig
    load_spots()        As LoadSpotConfig

    ' Schedules
    arrivals()          As ScheduledArrival
    shipments()         As ScheduledShipment

    ' Active queues
    active_unloads()    As ActiveUnload
    active_loads()      As ActiveLoad

    ' Pending queues (loads waiting for a spot)
    pending_unload_bbl() As Double
    pending_unload_mat() As String
    pending_load_bbl()  As Double
    pending_load_prod() As String

    ' Results
    snapshots()         As StepSnapshot

    ' Counts (UBound helpers since UDT arrays can't be dynamic without ReDim)
    num_raw_tanks       As Long
    num_blend_tanks     As Long
    num_recipes         As Long
    num_units           As Long
    num_product_tanks   As Long
    num_unload_spots    As Long
    num_load_spots      As Long
    num_arrivals        As Long
    num_shipments       As Long
    num_active_unloads  As Long
    num_active_loads    As Long
    num_pending_unloads As Long
    num_pending_loads   As Long
End Type
