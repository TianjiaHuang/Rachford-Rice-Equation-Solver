# Rachford–Rice Solver (MATLAB)

This repository provides MATLAB implementations of efficient and robust solvers for the multicomponent, multiphase **Rachford–Rice (RR) equations**, widely used in phase equilibrium calculations.

## References

1. Ryo Okuno, Russell T. Johns, and Kamy Sepehrnoori (2010).
   *A New Algorithm for Rachford–Rice for Multiphase Compositional Simulation.*
   SPE Journal, 15(02): 313–325.
   [https://doi.org/10.2118/117752-pa](https://doi.org/10.2118/117752-pa)

2. Huang, Johns, and Dindoruk (2026).
   *A Fast and Robust Convergent Algorithm for Rachford–Rice Equations and Its Extension to Negative Feed Compositions.*
   SPE Journal

## Features

* Newton-based RR solvers with line search for enhanced robustness
* Applicable to an arbitrary number of phases and components
* Designed for systems with positive overall compositions
* Improved convergence through advanced initialization and feasible-region control

## Files

*(See Huang et al. (2026) for detailed descriptions.)*

* `RR_Okuno.m` – RR solver based on Okuno et al. (2010)
* `RR_Huang.m` – Improved RR solver with a reduced feasible solution window
* `Initial_Estimation_*.m` – Initial guess strategies
* `main.m` – Example script demonstrating usage
* `RR_Huang.xlsx` – Example Excel File of the RR_Solver

## Usage

Run the example script:

```matlab
main
```

Or call the solver directly:

```matlab
beta = RR_Huang(z, K, beta0, tol);
```

### Inputs

* `z` : overall composition vector (1 × NC)
* `K` : equilibrium ratios ((NP − 1) × NC)
* `beta0` : initial estimate of phase fractions
* `tol` : convergence tolerance

### Output

* `beta` : computed phase mole fractions

Iteration history and convergence information are printed to the console.

## Notes

* The performance and robustness depend on the quality of the initial estimate.
* The reduced feasible-region (“smaller window”) approach improves convergence, especially near phase boundaries.
* The method is suitable for large-scale compositional simulations.
