# ExcelMonteCarlo

## Introduction

Monte-Carlo analysis is a common technique used within science and mathematics. Within my career I have frequently used it to help understand the effect of uncertainty and risk within cost and time of large procurement projects.

The purpose of this add-in is to provide a way to undertake simple monte carlo analysis without relying on complex and costly software while still having the basic functionality needed.

## Implementation

The tool is written in VBA as a .xlam add-in within Excel. While not the most performant of implementations it does offer a simple interface and an implementation of the common distributions used within the project management sphere.

Correlation has been implemented into the tool to allow the random sampling of correlated distributions. The implementation is based on Iman and Conover's 1982 paper.[^fn] The tool uses a Cholesky decomposition of the correlation matrix to check for consistency and then proceeds with Iman and Conover's method if consistent.

## Functionality

I have mirrored the functionality of @Risk so if you have used that before it should be familiar. The tool has the following functions implemented in Excel.

There are two types of function **Input** and **Output** the input functions generate Monte Carlo data for the relevant distribution which is stored in memory and the output functions perform that operation on the underlying slice data.

### Input Functions

#### MonteCarloInputTriang

This function Generates a [triangular distribution](https://en.wikipedia.org/wiki/Triangular_distribution) with given minimum, most likely and max.

#### MonteCarloInputBernoulli

Generates a [Bernoulli distribution](https://en.wikipedia.org/wiki/Bernoulli_distribution) with a given probability.

### Output Functions

#### MonteCarloOutputAverage

Calculates the arithmetic mean of the given distribution

#### MonteCarloOutputPercentile

Returns the percentile of the given distribution

### Installation

To install the add-in follow these ["Add or Remove COM add-in"](https://support.microsoft.com/en-gb/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) instructions.

You may need to unblock the add-in depending on how you downloaded it as shown in [these instructions](https://www.calctopia.com/unblock-excel-add-in/)

#### References

[^fn]: Iman, R. L., and W. J. Conover. 1982. "A Distribution-Free Approach to Inducing Rank Correlation Among Input Variables." Commun. Statist.-Simula. Computa. 11: 311-334. [link](https://www.uio.no/studier/emner/matnat/math/STK4400/v05/undervisningsmateriale/A%20distribution-free%20approach%20to%20rank%20correlation.pdf)
