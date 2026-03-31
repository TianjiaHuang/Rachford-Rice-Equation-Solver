Rachford-Rice Solver (Excel)

These codes provides Excel implementations of efficient and robust solvers for the multicomponent, multiphase Rachford–Rice (RR) equations, commonly used in phase equilibrium calculations.

References:
   1. Okuno, Johns, and Sepehrnoori. 2010. A new algorithm for Rachford-Rice for multiphase compositional simulation. SPE J  15 (02): 313-325. https://doi.org/10.2118/117752-pa
   2. Huang, Johns, and Dindoruk. 2026. A Fast and Robust Convergent Algorithm for Rachford-Rice Equations and Its Extension to Negative Feed Compositions. SPE J

Files:
(See Huang et al. (2026) for detailed descriptions.)
   1. RR_Huang.xlsx – Excel implementation of the RR solver without VBA macros.
   2. RR_Huang.xlsm – Excel implementation with VBA support.
   3. Initial_Estimation_Gradient_Huang.bas – VBA module for gradient-based initial estimation.
   4. RR_Huang_dynamic.bas – VBA module implementing the RR solver.

Usage Instructions:
   Option 1: Import VBA Modules into Excel (.xlsx)
   1. Open the Excel file.
   2. Press ALT + F11 to open the VBA editor.
   3. Navigate to: File → Import File.
   4. Import all .bas files.
   
   Option 2: Use Pre-configured Macro File (.xlsm)
   1. Before opening the .xlsm file, right-click the file.
   2. Select Properties → General.
   3. Click “Unblock” at the bottom.
   4. Open the file and enable macros when prompted.

Where:
   z : overall composition (1 × NC)
   K : equilibrium ratios ((NP−1) × NC)
   beta0 : initial estimate of phase fractions
   tol : convergence tolerance
   
   beta : computed phase mole fractions