# -*- coding: utf-8 -*-
"""
Created on Wed Mar 24 16:45:28 2021

@author: Administrator
"""
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Polygon


# Fixing random state for reproducibility
np.random.seed(19680801)

# fake up some more data
spread = np.random.rand(50) * 100
center = np.ones(25) * 40
flier_high = np.random.rand(10) * 100 + 100
flier_low = np.random.rand(10) * -100
d2 = np.concatenate((spread, center, flier_high, flier_low))
# Making a 2-D array only works if all the columns are the
# same length.  If they are not, then use a list instead.
# This is actually more efficient because boxplot converts
# a 2-D array into a list of vectors internally anyway.
data = [d2, d2[::2]]

labels = list('AB')

# Multiple box plots on one Axes
fig, ax = plt.subplots()
ax.boxplot(data,0, '', labels=labels)



x = np.array([0.80,1.2])
y = np.array([150,150])
#有了 x 和 y 数据之后，我们通过 plt.plot(x, y) 来画出图形，并通过 plt.show() 来显示。
ax.plot(x,y)

x = np.array([0.80,1.2])
y = np.array([-150,-150])

ax.plot(x,y)


plt.show()