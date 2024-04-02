# Solidworks-Macros

A collection of various Solidworks VBA scripts.

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

### Subroutine - mirrorCoordsX - Create copies of selected coordinate systems mirrrored on X-axis

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

$$
\begin{aligned}
\phi_Z &= \begin{cases} 0 & \text{if } r_{11} = r_{21} = 0 \\
            -\arctan2(-r_{21}, r_{11}) & \text{otherwise} \end{cases}\\

\phi_Y &= \begin{cases} -\pi / 2 & \text{if } r_{11} = r_{21} = 0 \\
            -\arctan2(r_{31}, \sqrt{r_{11}^2 + (-r_{21})^2}) & \text{otherwise} \end{cases}\\

\phi_X &= \begin{cases} -\arctan2(-r_{12}, r_{22}) & \text{if } r_{11} = r_{21} = 0 \\
            -\arctan2(r_{32}, r_{33}) & \text{otherwise} \end{cases}
\end{aligned}
$$

Usage: Select coordinate system features you want to copy, then run the script as instructed in the beginning.