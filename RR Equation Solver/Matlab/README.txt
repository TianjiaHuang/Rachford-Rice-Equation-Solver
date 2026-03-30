Rachford-Rice Solver (MATLAB)

These codes provides MATLAB implementations of efficient and robust solvers for the multicomponent, multiphase Rachford–Rice (RR) equations, commonly used in phase equilibrium calculations.

References:
   1. Okuno, Johns, and Sepehrnoori. 2010. A new algorithm for Rachford-Rice for multiphase compositional simulation. SPE J  15 (02): 313-325. https://doi.org/10.2118/117752-pa
   2. Huang, Johns, and Dindoruk. 2026. A Fast and Robust Convergent Algorithm for Rachford-Rice Equations and Its Extension to Negative Feed Compositions. SPE J

Features:
   1. Newton-based RR solvers with line search for enhanced robustness
   2. Applicable to an arbitrary number of phases and components
   3. Designed for systems with positive overall compositions
   4. Improved convergence through advanced initialization and feasible-region control

Files:
(See Huang et al. (2026) for detailed descriptions.)
   1. RR_Okuno.m – RR solver based on Okuno et al. (2010)
   2. RR_Huang.m – Improved RR solver with a reduced feasible solution window
   3. Initial_Estimation_*.m – Initial guess strategies
   4. main.m – Example script demonstrating usage

Usage:
   1. Run the example script: main.m
   2. Or call the solver directly:
      beta = RR_Huang(z, K, beta0, tol);

Where:
   z : overall composition (1 × NC)
   K : equilibrium ratios ((NP−1) × NC)
   beta0 : initial estimate of phase fractions
   tol : convergence tolerance
   
   beta : computed phase mole fractions
   Iteration history and convergence information are printed to the console.