'This program is meant to generate 4D plots that show localization/distribution of defects (pores) within CT-scanned components'

from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import pandas
import numpy as np
import os

fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')

os.chdir(r'C:\Users\Angu312\Data')
data = pandas.read_excel("Porosity_Report.xlsx", "Data")

X = np.array(data['Center X (mm)'].values.tolist())
Y = np.array(data['Center Y (mm)'].values.tolist())
Z = np.array(data['Center Z (mm)'].values.tolist())
pores = np.array(data['Diameter (μm)'].values.tolist())

group1 = (pores < 200)
group2 = ((pores >= 200) & (pores < 500))
group3 = ((pores >= 500) & (pores < 1000))
group4 = (pores >= 1000)

groups = [group1, group2, group3, group4]
colors = ['r', 'b', 'g', 'y'] # color order of groups

for group, color in zip(groups, colors):
    ax.scatter3D(X[group], Y[group], Z[group], c=color)

bounds = [0, 200, 500, 1000, 10000]
cmap = mcolors.ListedColormap(colors)
norm = mcolors.BoundaryNorm(bounds, len(colors))
scalebar = ax.scatter(X,Y,Z, c=pores, marker="o", cmap=cmap, norm=norm)
fig.colorbar(scalebar)

ax.set_title('Pore Size (μm) Distribution of CT-scanned Nozzle', fontsize=10)
ax.set_xlabel('Center X (mm)')
ax.set_ylabel('Center Y (mm)')
ax.set_zlabel('Center Z (mm)')

plt.show()
