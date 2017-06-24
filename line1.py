'''
Created on 2016-12-19

@author: sy
'''

import numpy as np
import matplotlib.pyplot as plt  # @UnresolvedImport
 
x = np.arange(-5.0, 5.0, 0.02)
y1 = np.sin(x)

plt.figure(1)
plt.subplot(211)
plt.plot(x, y1)
 
plt.subplot(212)
xlim=(-2.5, 2.5)
ylim=(-1, 1)
plt.plot(x, y1)
plt.show()