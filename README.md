# Campagna

## Introduction

This is an Excel add-in written in C# using the [Excel-Dna](https://excel-dna.net/) library that enables the running of monte-carlo analysis within Excel.

It was designed to give robust excel monte carlo analysis without using VBA and without too much complication.

## Initial Experimentation within VBA

Before going down the route of programming the tool in C# I generated a prototype within Excel using VBA. This allowed me to investigate the solution within Excel and also highlighted the issues that would result in the move to C#.

Within the prototype I managed to implement all of the features that would make it into Campagna and also some more advances features that could be ported across in the future. Specifically the implementation of correlated distributions.

Correlation was implemented into the tool to allow the random sampling of correlated distributions. The implementation is based on Iman and Conover's 1982 paper.[^fn] The tool uses a Cholesky decomposition of the correlation matrix to check for consistency and then proceeds with Iman and Conover's method if consistent.

The add-in can be found in the [ExcelMonteCarlo](./ExcelMonteCarlo/) folder in this repo.

### Issues with VBA

While VBA offered a quick way to access Excel and develop solutions for it there were a number of failings that caused me to investigate other possible ways to implement the tool:

- **Speed** - VBA is not performant when dealing with large in memory calculations and does not have a lot of the newer programming structures that allow for easier computation. While I didn't implement multithreading in the final tool C# would allow this and it could be implemented in future while it would not be possible in VBA.

- **Version Control** - As anyone who has worked with VBA for any length of time knows, version control is very hard to do and can lead to real issues with deployment and issue tracking. Moving to C# would allow the use of a modern VCS such as Git to combat this.

- **Code maintainability and deployment** - With VBA you are limited in options for IDE and also for collaboration. The move to C# would allow easier code generation within a modern IDE and easier sharing and deployment of final versions.

## Move to C\#

The move to C# was made possible by the brilliant Excel-Dna library. This is an open source library which provides and API to interact with Excel Objects built upon a C framework to allow for very quick execution while being able to use the more high-level languages VB.NET, F# and C#.

## Functionality

I have mirrored the functionality of @Risk so if you have used that before it should be familiar. The tool has the following functions implemented in Excel.

There are two types of function **Input** and **Output** the input functions generate monte carlo data for the relevant distribution which is stored in memory and the output functions perform that operation on the underlying slice data.

### Input Functions

#### CampagnaInputTriangular
This function Generates a [triangular distribution](https://en.wikipedia.org/wiki/Triangular_distribution) with given minimum, most likely and max.

#### CampagnaInputBernoulli
Generates a [Bernoulli distribution](https://en.wikipedia.org/wiki/Bernoulli_distribution) with a given probability.

### Output Functions

#### CampagnaOutputSingleSlice
Returns the kth data slice of the distribution in order that they were sampled. This is useful for debugging.

#### CampagnaOutputMean
Calculates the arithmetic mean of the given distribution

#### CampagnaOutputPercentile
Returns the percentile of the given distribution

### Model Run

Once the model has been set up in Excel the tool can be run by pressing the Generate Results button within the Campagna Window. A seed value integer can be provided to allow for repeatable results. It can handle any C# integer value. \[-2,147,483,648 to 2,147,483,647\]

| ![Campagna Ribbon](./Documentation/Campagna%20Ribbon.png) |
| :--: |
| Campagna Ribbon |

## Installation

Place the following files in the same folder:

### 32-bit

[Campagna-AddIn-Packed.xll](./bin/Release/Campagna-AddIn-packed.xll)

[ExcelDna.Intellisense.dll](./bin/Release/ExcelDna.IntelliSense.dll)

### 64-bit

[Campagna-AddIn64-Packed.xll](./bin/Release/Campagna-AddIn64-packed.xll)

[ExcelDna.Intellisense.dll](./bin/Release/ExcelDna.IntelliSense.dll)

### In Excel

Then add the add-in to excel by following the ["Add or Remove COM add-in"](https://support.microsoft.com/en-gb/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) instructions at the included link.

You may need to unblock the add-in depending on how you downloaded it as shown in [these instructions](https://www.calctopia.com/unblock-excel-add-in/)


[^fn]: Iman, R. L., and W. J. Conover. 1982. "A Distribution-Free Approach to Inducing Rank Correlation Among Input Variables." Commun. Statist.-Simula. Computa. 11: 311-334. [link](https://www.uio.no/studier/emner/matnat/math/STK4400/v05/undervisningsmateriale/A%20distribution-free%20approach%20to%20rank%20correlation.pdf)