# Monte Carlo Integration in Excel/VBA

## Overview

This project implements **Monte Carlo numerical integration in Microsoft Excel using VBA**.  
It was developed as a university project to combine probability theory, numerical methods, object-oriented programming, and data visualization in one working analytical tool.

The workbook can approximate definite integrals over three-dimensional rectangular domains, compare Monte Carlo estimates with exact values for selected test functions, analyse estimation error, and automatically generate charts illustrating convergence.

## Project goals

The project was designed to:

- implement Monte Carlo integration for functions defined on a domain of the form  
  \([a,b] \times [c,d] \times [e,f]\),
- illustrate the convergence of Monte Carlo estimators,
- analyse the distribution and magnitude of estimation errors,
- implement the computational logic in **VBA**,
- use **object-oriented programming** to organise the calculation engine,
- automate calculations, error analysis, statistics, and chart generation inside Excel.

## How it works

For a function \(f\), the Monte Carlo estimator approximates an integral by evaluating the function at randomly generated points and averaging the results.

As the number of simulations increases, the estimator should converge toward the true value of the integral. The project demonstrates this behaviour experimentally and measures the corresponding estimation error.

The workbook contains three test functions (`f1`, `f2`, `f3`) used to compare:

- Monte Carlo estimates,
- exact integral values,
- estimation errors,
- estimator behaviour for different sample sizes.

## Main functionality

### Monte Carlo engine

The VBA implementation:

- generates pseudorandom points inside a specified integration domain,
- evaluates the selected function at those points,
- calculates the Monte Carlo approximation,
- supports configurable integration limits,
- resets and reinitialises calculations when parameters change.

### Error analysis

The program compares the simulated estimate with the exact value of the integral and calculates the estimation error.

Repeated simulations are used to study:

- estimator stability,
- error magnitude,
- convergence as the sample size increases,
- distributions of estimation results.

### Automated visualisation

The workbook automatically creates charts based on calculated results, including:

- convergence plots,
- estimation-error plots,
- histograms of estimators for different sample sizes,
- box plots illustrating estimator distributions.

## VBA structure

The project uses both standard VBA modules and a class-based implementation.

### Main procedures and functions

Examples of implemented functionality include:

- `Calka_MC` – performs repeated Monte Carlo simulations,
- `Calka1` – performs a single Monte Carlo update,
- `PoliczBlad` – calculates the estimation error,
- `TworzenieEstymatorow` – creates estimator samples and calculates statistics,
- `TworzenieWykresuZTablicy` – automatically generates charts from calculated data,
- `Funkcja_na_0_1` – evaluates repeated simulations on the unit interval/domain,
- `funkcja_na_innych_przedzialach` – performs calculations for configurable ranges,
- `Blad_na_0_1` and `Blad_na_innych_przedzialach` – generate error samples,
- `CaleZadanie` – runs the complete calculation and visualisation workflow.

The VBA class stores the integration parameters, selected function, current estimate, sample count, and domain volume, and exposes methods for changing functions and integration ranges.

## Example findings

The simulations show the expected Monte Carlo behaviour:

- with small sample sizes, estimates fluctuate considerably,
- estimation errors are larger and more dispersed,
- as the number of simulations increases, estimates become more stable,
- the Monte Carlo approximation converges toward the exact integral value.

The generated charts make this behaviour visible directly in Excel.

## Technologies

- **Microsoft Excel**
- **VBA**
- **Object-Oriented Programming in VBA**
- **Monte Carlo methods**
- **Probability and statistics**
- **Numerical integration**
- **Automated data visualisation**

## What I learned

This project gave me practical experience in:

- translating a mathematical method into a working analytical tool,
- structuring calculation logic in VBA,
- using object-oriented programming in Excel,
- automating repetitive analytical tasks,
- validating numerical results against reference values,
- analysing estimator error and convergence,
- generating visualisations programmatically,
- working collaboratively on a quantitative programming project.

## Authors

- Kacper Mocarski
- Krzysztof Michalak
- Łukasz Trojgo
