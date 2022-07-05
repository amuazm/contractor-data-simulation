import numpy as np

n, k = 329492, 5
vals = np.random.default_rng().dirichlet(np.ones(k), size=1)
k_nums = [round(v) for v in vals[0]*n]

print(k_nums)