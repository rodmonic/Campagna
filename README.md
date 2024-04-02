# Campagna

## Introduction

This is an Excel add-in written in C# using the [excel-dna](https://excel-dna.net/) library that enables the running of monte-carlo analysis within Excel. 

It was designed to give robust excel monte carlo analysis without using VBA and without too much complication.

## Functionality

I have mirrored the fucntionaility of @Risk so if you have used that before it should be familiar. The tool has the following functions implemented in Excel.

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

Once the modeel has been set up in Excel the tool can be run by pressing the Generate Results button within the Campagna Window. A seed value integer can be provided to allow for repeatable results. It can handle any C# integer value. [-2,147,483,648 to 2,147,483,647]

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

Then add the add-in to excel by following the ["Add or Remove COM add-in"](https://support.microsoft.com/en-gb/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) instructions.

You may need to unblock the add-in depending on how you downloaded it as shown in [these instructions](https://www.calctopia.com/unblock-excel-add-in/)
