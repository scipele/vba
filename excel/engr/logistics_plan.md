# Plant Logistics Inventory Simulation — Program Plan

## Overview
A VBA-based time-step simulation that models raw material receiving, blending,
unit processing, product storage, and product shipping at a plant. The user
configures the model via Excel tables (ListObjects) on input sheets, runs the
simulation, and reviews results in an output table plus an optional UserForm
with tank/railcar graphics.

---

## 1. Excel Workbook Layout

### Sheet: "Config"
| Table Name         | Purpose                                          | Key Columns |
|--------------------|--------------------------------------------------|-------------|
| `tblRunConfig`     | Global run parameters                            | ParamName, ParamValue |

**tblRunConfig rows:**
| ParamName               | ParamValue (example) | Notes                              |
|-------------------------|----------------------|------------------------------------|
| RunDuration_Days        | 30                   | Total simulation length            |
| TimeStep_Hours          | 1                    | Granularity (1 = hourly)           |
| UnloadOnWeekends        | FALSE                | TRUE/FALSE                         |
| LoadOnWeekends          | FALSE                | TRUE/FALSE                         |
| StartDate               | 2026-04-01           | Calendar anchor                    |

### Sheet: "RawMaterials"
| Table Name              | Purpose                                      | Key Columns |
|-------------------------|----------------------------------------------|-------------|
| `tblUnloadSchedule`     | Arrival calendar for inbound shipments       | ArrivalDay, Mode, Quantity_BBL, MaterialName |
| `tblUnloadSpots`        | Unloading infrastructure per mode            | Mode, NumSpots, AvgUnloadTime_Hrs, BBLperLoad |
| `tblRawTanks`           | Raw-material tank farm                       | TankName, MaterialName, Capacity_BBL, StartInventory_BBL, MinInventory_BBL |

- **Mode** values: Rail, Truck, Barge
- **ArrivalDay**: integer day of simulation (1-based) when loads arrive
- **AvgUnloadTime_Hrs**: hours to unload one load (railcar, truck, barge)
- **BBLperLoad**: volume per single load unit

### Sheet: "Blending"
| Table Name           | Purpose                                        | Key Columns |
|----------------------|------------------------------------------------|-------------|
| `tblBlendTanks`      | Blend tank definitions                         | BlendTankName, Capacity_BBL, StartInventory_BBL |
| `tblBlendRecipes`    | Material proportions per blend                 | BlendTankName, MaterialName, FractionOfBlend |

### Sheet: "Processing"
| Table Name         | Purpose                                          | Key Columns |
|--------------------|--------------------------------------------------|-------------|
| `tblUnits`         | Processing unit definitions                      | UnitName, DesignCapacity_BBL_Day, FeedSource, ProductName |

- **FeedSource**: name of a raw tank or blend tank that feeds the unit

### Sheet: "Products"
| Table Name            | Purpose                                       | Key Columns |
|-----------------------|-----------------------------------------------|-------------|
| `tblProductTanks`     | Product storage tank farm                     | TankName, ProductName, Capacity_BBL, StartInventory_BBL, MinInventory_BBL |
| `tblLoadSchedule`     | Outbound product shipping (rail loading)      | ShipDay, ProductName, Quantity_BBL, Mode |
| `tblLoadSpots`        | Loading infrastructure                        | Mode, NumSpots, AvgLoadTime_Hrs, BBLperLoad |

### Sheet: "Results"
| Table Name         | Purpose                                          |
|--------------------|--------------------------------------------------|
| `tblSimResults`    | One row per time step with all tank inventories, throughputs, flags |

**tblSimResults columns (auto-generated):**
SimStep, DateTime, [each raw tank inventory], [each blend tank inventory],
[each product tank inventory], [each unit throughput], UnloadingActive,
LoadingActive, Flags

---

## 2. VBA Module Architecture

| Module / Item           | Type       | Responsibility                                           |
|-------------------------|------------|----------------------------------------------------------|
| `modMain`               | Standard   | Entry point `RunSimulation`, orchestration               |
| `modSetup`              | Standard   | `SetupInputTables` — creates sheets & ListObjects        |
| `modSimEngine`          | Standard   | Core time-step loop, inventory math                      |
| `modResults`            | Standard   | Write results to tblSimResults, summary stats            |
| `modHelpers`            | Standard   | Table lookups, date helpers, validation                  |
| `frmDashboard`          | UserForm   | Graphical tank display, time-step slider                 |

---

## 3. Data Structures (UDTs)

```
RawTank        — name, material, capacity, current inventory, min
BlendTank      — name, capacity, current inventory, recipe array
ProcessingUnit — name, capacity (BBL/hr), feed source, product name
ProductTank    — name, product, capacity, current inventory, min
UnloadSpot     — mode, count, avg time, queue of pending loads
LoadSpot       — mode, count, avg time, queue of pending shipments
SimState       — arrays of above, current step, flags
```

**Multi-tank support:**  Multiple raw tanks can hold the same material; multiple
product tanks can hold the same product.  The simulation cascades
deposit/withdraw operations across all matching tanks:
- `DepositToRawTanks` — fills tanks in order; overflows to next tank
- `WithdrawFromRawTanks` — drains tanks in order; moves to next
- `DepositToProductTanks` — same cascade for products
- `WithdrawFromProductTanks` — same cascade for products
- `TotalRawInventoryByMaterial` / `TotalProductInventoryByProduct` — sums
  across all matching tanks for availability checks
- `FindAllRawTanksByMaterial` / `FindAllProductTanksByProduct` — returns all
  matching tank indices

---

## 4. Simulation Loop (Pseudo-code)

```
For each time step (hour):
    1. Check if arrivals are scheduled → add to unload queue
    2. Process unloading queue (respect # spots, unload time, weekend flag)
       → increment raw tank inventories (cap at capacity, flag overflow)
    3. Blend (if applicable): pull from raw tanks per recipe fractions
       → increment blend tank inventories
    4. Process units: pull from feed source at hourly rate
       → decrement feed tank, increment product tank
       → flag if feed tank < min or product tank > capacity
    5. Check if shipments scheduled → add to load queue
    6. Process loading queue → decrement product tank inventories
       → flag if product tank < min
    7. Record all inventories & flags to results array
Next step
Write results array → tblSimResults
```

---

## 5. UserForm Graphics Concept

```
┌─────────────────────────────────────────────────────┐
│  [<<] Step: 47 / 720  [>>]   Date: 2026-04-02 23:00│
│                                                     │
│  ══ INBOUND ══         ══ RAW TANKS ══              │
│  ┌───┐ ┌───┐          ╭─────╮  ╭─────╮             │
│  │R/C│ │R/C│   ───►   │ TK1 │  │ TK2 │             │
│  └───┘ └───┘          │ 72% │  │ 45% │             │
│                        ╰─────╯  ╰─────╯             │
│                            │                        │
│  ══ BLEND TANKS ══        ▼                         │
│  ╭─────╮           ══ UNITS ══                      │
│  │ BL1 │  ◄─────   [UNIT-1 @95%]                   │
│  │ 60% │                │                           │
│  ╰─────╯                ▼                           │
│                   ══ PRODUCT TANKS ══               │
│                   ╭─────╮  ╭─────╮                  │
│                   │ PT1 │  │ PT2 │   ───►  ┌───┐   │
│                   │ 88% │  │ 33% │         │R/C│   │
│                   ╰─────╯  ╰─────╯         └───┘   │
│                                      ══ OUTBOUND ══ │
│  [Run]  [Reset]  [Export Chart]                     │
└─────────────────────────────────────────────────────┘
```

- Tanks drawn as rounded rectangles with fill level shading
- Railcars/trucks as small rectangles
- Color coding: green (normal), yellow (low/high warning), red (violation)
- Slider or step buttons to scrub through time

---

## 6. Implementation Phases

| Phase | Scope                                             |
|-------|---------------------------------------------------|
| 1     | `modSetup` — create all sheets & input tables     |
| 2     | UDTs + `modHelpers` — data loading from tables    |
| 3     | `modSimEngine` — core loop, no blending first     |
| 4     | Add blending logic                                |
| 5     | `modResults` — write output table + summary       |
| 6     | `frmDashboard` — UserForm graphics                |
| 7     | Testing, edge cases, polish                       |
