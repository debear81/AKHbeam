# Macaulay Beam Analysis for Excel (AKHbeam)
AKHbeam is a Microsoft Excel add-in (XLAM) written in VBA that provides worksheet functions for one-dimensional beam analysis using the Macaulay (double integration) method.

All functions utilize a common calculation engine and input format, allowing reactions, shear, moment, slope, and deflection results to be generated directly within Excel.

      = V_Beam[Function] (Beam Segments Range, Supports Range, Distributed Loads Range, Point Loads Range)

      Example: =V_BeamAll(A2:D5,F2:J4,L2:Q8,S2:U5)

      where:  A2:D5 = Beam Segments Range (r rows x 4 columns)
              F2:J4 = Supports Range (r rows x 5 columns)
              L2:Q8 = Distributed Loads Range (r rows x 6 columns)
              S2:U5 = Point Loads Range (r rows x 4 columns)

## FEATURES
- Single & multi-span beam analysis
- Point loads
- Uniform and linearly varying distributed loads
- Applied moments
- Support settlements
- Variable member properties (E and I)
- Shear, moment, slope, and deflection output
- Automatic reaction calculations
- Analysis result point output suitable for plotting and graph generation

## INSTALLATION
1. After downloading and saving the XLAM, open Excel. 
2. Click File --> Options. This will open the "Excel Options" dialog box.
3. On the lefthand menu, click "Add-ins". Near the bottom, next to the "Manage [Excel Add-ins]" label, click "Go...". This will open the "Add-ins" dialog.
4. Click "Browse..." and locate / select the XLAM file.

The add-in inludes a ribbon tab with help and information regarding the use and capabilities of the add-in. Once loaded in the Excel environment, the tab should be visible.

<p><img width="500" alt="ribbon-tab" src="https://github.com/user-attachments/assets/5d81bdfe-b1f4-44ea-87ff-3b7144f52ac2" /></p>

## WORKING UNITS
Functions are NOT unit-aware.

Input user ranges/values must be of a consistent unit system.

## FUNCTION DESCRIPTION
The add-in contains four UDF functions for output, which all rely on the same background calculation engine.

| Function      | Description               |
| ------------  | ------------              |
| V_BeamAll     | Detailed analysis results |
| V_BeamReac    | Support reactions         |
| V_BeamEq      | Equilibrium check         |
| V_BeamMaxMin  | Extreme values summary    |

Descriptions of the output ranges (tables) are as follows:
| V_BeamAll         | will return                                            |
| ----------        | ------------                                           |
| Position (X)      | position along the beam                                |
| Area (A)          | cross sectional area of beam                           |
| Shear (V)         | resulting internal shear in beam at X                  |
| Moment (M)        | resulting internal moment in beam at X                 |
| Stiffness (EI)    | member stiffness (E * I) at X                          |
| RAW Slope         | intermediate deflection value used during calculations |
| RAW Deflection    | intermediate deflection value used during calculations |
| Slope (theta)     | final calculated beam slope at X                       |
| Delection (delta) | final calculated beam deflection at X                  | 

| V_BeamReac        | will return |
| ----------        | ------------ |
| Position (X)      | position along the beam |
| Imp. disp. (dy)   | imposed displacement (settlement) of support, prior to calculations |
| Imp. rot. (theta) | imposed rotation of support prior to calculations |
| Vert. Reaction (Ry) | vertical reaction at support |
| Mom. Reaction (Mz) | moment reaction at support |

| V_BeamEq   will return |
| -----------------------|
| Load summations and the reaction summations. This is intended as a general QA/QC function. |

| V_BeamMaxMin | will return |
| ----------        | ------------ |
| Item | Max +V, Max -V, Max +M, Max -M, or Max Deflection |
| Position (X) | position along the beam where the max/min occurs |
| Shear (+/- V) | max/min resulting shear force in beam |
| Moment (+/- M) | max/min resulting moment force in beam |
| Deflection (delta) | absolute maximum deflection of beam |

## CAPBILITIES & ASSUMPTIONS
- Single & multi-span beam analysis
- Linear elastic analysis
- Small displacement theory
- Euler-Bernoulli beam formulation
- Variable E and I permitted by segment
- Pin supports assumed
- Support settlements permitted
- Applied moments permitted
- Units are user-defined and must be internally consistent

## LIMITATIONS
Current version limitations include:
- No fixed, roller, or spring support
- No Timoshenko (shear deformation) analysis
- No dynamic analysis
- No (automatic) moving load analysis
- No influence line generation
- No geometric nonlinearity
- No material nonlinearity
- No frame or truss analysis

### USER INPUT RANGES
<p><img width="500" alt="image_inputranges" src="https://github.com/user-attachments/assets/a9419ae9-8626-4df6-88af-7b7210286cd5" /></p>

### OUTPUT RANGES
<p><img width="500" alt="image_outputranges" src="https://github.com/user-attachments/assets/1eef50a7-6546-4fc5-a6a3-3a0cc05c5e3b" /></p>  

### EXAMPLES OF POSSIBLE USER-DEFINED GRAPHS
<p><img width="450" alt="image_graphs" src="https://github.com/user-attachments/assets/5abcc26a-5029-47fe-b38c-6e696617354f" /></p>

## VERIFICATION
See documents folder for comparisons of add-in results to industry software results

## ACKNOWLEDGEMENT
This project was originally inspired by Doug Jenkins’ Newton Excel Bach post,“Beam Actions and Deflections by Macaulay’s Method,” which demonstrated the se of Macaulay functions for beam reactions, shear, moment, slope, and deflection calculations in Excel.

[“Beam Actions and Deflections by Macaulay’s Method”](https://newtonexcelbach.com/2011/11/12/beam-actions-and-deflections-by-macaulays-method/)

AKHbeam’s code structure, worksheet functions, validation, ribbon interface, and add-in implementation were developed independently for this project.



