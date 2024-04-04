# Solidworks-Macros

A collection of various Solidworks VBA scripts, macros, functions, subroutines and other bits & pieces.

## Usage

1. Open Solidworks
2. From the ribbon menu at the top, select **Tools -> Macro -> New...**
3. Select a save location for the macro file
4. Copy the contents of desired macro file and paste them to your new macro file
5. Run the macro from **Tools -> Macro -> Run...** or press play from the VBA editor

## Included macros

### Function - Atan2 - two-argument arctangent

The function for computing the Atan2 is included in `atan2.vba`

Input: `y As Double, x As Double`  
Output: `Double`

$$
\text{atan2}(y, x) =
\begin{cases}
    \arctan\left(\frac{y}{x}\right) & \text{if } x > 0 \\
    \arctan\left(\frac{y}{x}\right) + \pi & \text{if } x < 0 \text{ and } y \geq 0 \\
    \arctan\left(\frac{y}{x}\right) - \pi & \text{if } x < 0 \text{ and } y < 0 \\
    +\frac{\pi}{2} & \text{if } x = 0 \text{ and } y > 0 \\
    -\frac{\pi}{2} & \text{if } x = 0 \text{ and } y < 0 \\
    \text{undefined} & \text{if } x = 0 \text{ and } y = 0
\end{cases}
$$

### Subroutine - mirrorCoordsX - Create X-axis-mirrored copies of selected coordinate systems

This subroutine creates a copy of each reference coordinate system mirrored about the X-axis (YZ plane) while preserving the right-handedness of the coordinate system.

> Use case: when designing a symmetrical sensor holder for multi-axis sensors using mirrored geometry, one has to mirror the sensor positions for the other side as well, while preserving the axes of the sensor coordinate system.

In Solidworks this is achieved by extracting the rotation matrix of the manually placed coordinate system, which is then converted to Tait-Bryan rotation angles for creating a duplicate coordinate system on the other side. The X-coordinates of the XYZ vectors (=first column of the rotation matrix $R$) are inverted to mirror about YZ plane. The X-vector is then inverted by inverting all of its components (=first row of the rotation matrix), hence the signs here do not match the ordinary formulation. Solidworks does not seem to rotate according to the right hand rule, hence the rotations are also inverted.

The Tait-Bryan **ZYX** angles are computed from the rotation matrix as

$$
R = 
\begin{bmatrix}
r_{11} & r_{12} & r_{13} \\
r_{21} & r_{22} & r_{23} \\
r_{31} & r_{32} & r_{33} \\
\end{bmatrix}
$$

```math
\begin{aligned}
\text{if } r_{11} = r_{21} = 0 \hspace{1em} & \begin{cases} \phi_Z = 0 & \\
                                                            \phi_Y = -\pi / 2 \\
                                                            \phi_X = -\arctan2(-r_{12}, r_{22})
                                              \end{cases} \\[10pt]

\text{otherwise} \hspace{1em} & \begin{cases} \phi_Z = -\arctan2(-r_{21}, r_{11}) \\
                                              \phi_Y = -\arctan2(r_{31}, \sqrt{r_{11}^2 + (-r_{21})^2}) \\
                                              \phi_X = -\arctan2(r_{32}, r_{33})
                                \end{cases}
\end{aligned}
```

**Usage:** Select coordinate system features you want to copy, then run the script as instructed in the beginning.


### Subroutine - saveCoordsToCsv - Save coordinate system locations and angles to csv file

This subroutine iterates over all selected coordinate systems and saves them to a csv file. The saved properties are the feature name, XYZ location and local coordinate axes X, Y and Z as unit vectors. As with all Solidworks API calls, the units are returned in meters.

An example of the output:
```csv
Slot,X,Y,Z,XX,XY,XZ,YX,YY,YZ,ZX,ZY,ZZ
Sensor coordinates B4,-0.0275,0.0802,0.0945,0.9379,0.1643,0.3055,-0.0101,0.8934,-0.4492,-0.3467,0.4182,0.8396
Sensor coordinates B3,0.0275,0.0802,0.0945,0.9379,-0.1643,-0.3055,0.0101,0.8934,-0.4492,0.3467,0.4182,0.8396
Sensor coordinates B2,-0.0268,0.0514,0.1040,0.9404,0.0870,0.3288,-0.0309,0.9846,-0.1721,-0.3387,0.1517,0.9286
```

**Usage:** Select coordinate system features you want to export, then run the script as instructed in the beginning.