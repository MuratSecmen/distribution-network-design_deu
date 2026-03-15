/*
 * Distribution Network Design and Optimization in a Supply Chain
 * ===============================================================
 * Published : Beykoz Akademi Dergisi, 2015; 3(1), 67-84
 * Authors   : Murat Secmen, Temel Oncan, Okan Tuna
 * Solver    : IBM ILOG CPLEX Optimization Studio (OPL)
 *
 * Problem:
 *   Multi-period, multi-modal supply chain network optimization.
 *   Minimize: transportation + opportunity + holding + stockout costs.
 *
 * Network:
 *   Suppliers : T1=Russia, T2=Germany, T3=Norway  (i in 1..3)
 *   Factories : F1=Spain, F2=Russia               (j in 1..2)
 *   Dist Ctr  : Ukraine (single DC)
 *   Markets   : M1=Finland, M2=Turkey, M3=Russia  (k in 1..3)
 *   Modes     : 1=Rail, 2=Road, 3=Air             (t in 1..3)
 *   Periods   : 6 months                          (p in 1..6)
 */

// ── INDEX SETS ───────────────────────────────────────────────────────────────
int nSuppliers  = 3;
int nFactories  = 2;
int nModes      = 3;
int nPeriods    = 6;
int nCustomers  = 3;

range Suppliers = 1..nSuppliers;
range Factories = 1..nFactories;
range Modes     = 1..nModes;
range Periods   = 1..nPeriods;
range Customers = 1..nCustomers;
range Periods0  = 0..nPeriods;   // includes period 0 for initial conditions

// ── PARAMETERS (read from .dat file) ─────────────────────────────────────────

// Transportation cost: supplier i -> factory j, mode t (period-independent)
float C[Suppliers][Factories][Modes] = ...;

// Factory-to-DC cost: factory j, period p
float Cw[Factories][Periods] = ...;

// DC-to-customer cost: customer k, period p
float Cy[Customers][Periods] = ...;

// Opportunity cost coefficient N_ijt (lead time in time units)
float N[Suppliers][Factories][Modes] = ...;

// Opportunity cost rate (monetary/time unit)
float pi_rate = ...;

// Supplier capacity alpha_ip
float alpha[Suppliers][Periods] = ...;

// Factory capacity b_jp
float b_cap[Factories][Periods] = ...;

// DC capacity c_p
float c_cap[Periods] = ...;

// Transport mode capacity A_tp
float A_mode[Modes][Periods] = ...;

// Customer demand d_kp
float d_demand[Customers][Periods] = ...;

// Inventory holding cost per unit at DC
float R = ...;

// Backlog/stockout cost per unit
float T_cost = ...;

// Initial DC inventory
float Q_init = ...;

// ── DECISION VARIABLES ───────────────────────────────────────────────────────

// X_ijtp: units from supplier i to factory j, mode t, period p
dvar float+ X[Suppliers][Factories][Modes][Periods];

// W_jp: units from factory j to DC in period p
dvar float+ W[Factories][Periods];

// Y_kp: units from DC to customer k in period p
dvar float+ Y[Customers][Periods];

// B_kp: backlogged demand of customer k in period p (p=0 means initial=0)
dvar float+ B[Customers][Periods0];

// Q_p: DC inventory at end of period p (p=0 = initial inventory)
dvar float+ Q[Periods0];

// ── OBJECTIVE FUNCTION ────────────────────────────────────────────────────────
minimize
  // (1) Transportation cost: supplier -> factory
  sum(i in Suppliers, j in Factories, t in Modes, p in Periods)
      C[i][j][t] * X[i][j][t][p]

  // (2) Factory -> DC cost
+ sum(j in Factories, p in Periods)
      Cw[j][p] * W[j][p]

  // (3) DC -> Customer cost
+ sum(k in Customers, p in Periods)
      Cy[k][p] * Y[k][p]

  // (4) Opportunity cost (time-based penalty for transport delays)
+ sum(i in Suppliers, j in Factories, t in Modes, p in Periods)
      N[i][j][t] * pi_rate * X[i][j][t][p]

  // (5) Inventory holding cost at DC
+ sum(p in Periods)
      R * (sum(j in Factories) W[j][p]
           - sum(k in Customers) Y[k][p]
           - sum(k in Customers) B[k][p]
           + Q[p-1])

  // (6) Stockout / backlog cost
+ sum(p in Periods)
      T_cost * (sum(k in Customers) Y[k][p]
                + sum(k in Customers) B[k][p]
                - sum(j in Factories) W[j][p]
                - Q[p-1]);

// ── CONSTRAINTS ──────────────────────────────────────────────────────────────
subject to {

  // C1: Supplier capacity
  forall(i in Suppliers, p in Periods)
    SupplierCap:
      sum(j in Factories, t in Modes) X[i][j][t][p] <= alpha[i][p];

  // C2: Factory capacity
  forall(j in Factories, p in Periods)
    FactoryCap:
      W[j][p] <= b_cap[j][p];

  // C3: DC capacity
  forall(p in Periods)
    DCCap:
      sum(k in Customers) Y[k][p] <= c_cap[p];

  // C4: Customer demand satisfaction (with backlog carry-over)
  forall(k in Customers, p in Periods)
    DemandSat:
      Y[k][p] + B[k][p-1] + B[k][p] >= d_demand[k][p];

  // C5: Transport mode capacity
  forall(t in Modes, p in Periods)
    ModeCap:
      sum(i in Suppliers, j in Factories) X[i][j][t][p] <= A_mode[t][p];

  // C6: Flow balance — factory inflow equals outflow to DC
  forall(j in Factories, p in Periods)
    FlowBalance:
      sum(i in Suppliers, t in Modes) X[i][j][t][p] == W[j][p];

  // C7: DC inventory balance
  InitInventory:
    Q[0] == Q_init;

  forall(p in Periods)
    InvBalance:
      Q[p] == Q[p-1]
             + sum(j in Factories) W[j][p]
             - sum(k in Customers) Y[k][p];

  // C8: Initial backlog = 0
  forall(k in Customers)
    InitBacklog:
      B[k][0] == 0;

}

// ── POST-PROCESSING ───────────────────────────────────────────────────────────
execute DISPLAY_RESULTS {
  writeln("==============================================");
  writeln("  OPTIMAL SOLUTION");
  writeln("  Objective Value (Min. Z) = " + cplex.getObjValue());
  writeln("==============================================");

  writeln("\n-- Supplier -> Factory Flows (X_ijtp, nonzero) --");
  var supplier_names = ["Russia(T1)", "Germany(T2)", "Norway(T3)"];
  var factory_names  = ["Spain(F1)", "Russia(F2)"];
  var mode_names     = ["Rail", "Road", "Air"];
  for(var i in Suppliers)
    for(var j in Factories)
      for(var t in Modes)
        for(var p in Periods)
          if(X[i][j][t][p] > 0.01)
            writeln("  X["+i+"]["+j+"]["+t+"]["+p+"] = "+X[i][j][t][p]
                    +" ("+supplier_names[i-1]+"->"+factory_names[j-1]
                    +" via "+mode_names[t-1]+", Period "+p+")");

  writeln("\n-- Factory -> DC Flows (W_jp, nonzero) --");
  for(var j in Factories)
    for(var p in Periods)
      if(W[j][p] > 0.01)
        writeln("  W["+j+"]["+p+"] = "+W[j][p]);

  writeln("\n-- DC -> Customer Flows (Y_kp, nonzero) --");
  var cust_names = ["Finland(M1)", "Turkey(M2)", "Russia(M3)"];
  for(var k in Customers)
    for(var p in Periods)
      if(Y[k][p] > 0.01)
        writeln("  Y["+k+"]["+p+"] = "+Y[k][p]+" (->"+cust_names[k-1]+")");

  writeln("\n-- DC Inventory (Q_p) --");
  for(var p in Periods0)
    if(Q[p] > 0.01)
      writeln("  Q["+p+"] = "+Q[p]);

  writeln("\n-- Backlog (B_kp, nonzero) --");
  for(var k in Customers)
    for(var p in Periods0)
      if(B[k][p] > 0.01)
        writeln("  B["+k+"]["+p+"] = "+B[k][p]);
}
