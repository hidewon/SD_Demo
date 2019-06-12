# %%

# まだ環境にJupyterがはいっていないのでRun Cellみただけで満足ってことで。
import math
import random

def pdf(x,mu,sigma):
    sigma_pow=sigma*2
    return (
        1
        /math.sqrt(2*math.pi*sigma_pow)
        *math.exp(-(x-mu)**2/2*sigma_pow)
    )

#%%
