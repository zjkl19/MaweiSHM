# -*- coding: utf-8 -*-

import numpy as np
import pandas as pd
 
import matplotlib.pyplot as plt
 
  
 
data = {
 
  '1#': [2.6,2,2.5,1.2,1.6,1.4,1.2,2,2,2,2],
  '2#': [0.8,0.6,1,0.1,0.6,0.6,0.5,1.3,1.3,1.3,1.4],
  '3#': [-1.7,-1.4,-2,-0.6,-1,-0.9,-0.7,-1.4,-1.4,-1.4,-1.5],
  "4#": [-0.8,-0.5,-1,-0.4,-0.9,-0.9,-0.7,-1.2,-1.2,-1.2,-1.2]
 
}
 
df = pd.DataFrame(data)
 
  
# df.plot.box(title="Consumer spending in each country", vert=False)
 
#df.plot.box(title="Consumer spending in each country")

df.plot.box()

plt.grid(linestyle="--", alpha=0.3)

plt.show()