"""Astronomical Coordinadte Converter For Interstellar Navigation"""
# ngc188_center = SkyCoord(12.11*u.deg, 85.26*u.deg, frame='icrs')
# ngc188_center
# from astropy.io import fits
# from astropy.table import QTable
# from astropy.utils.data import download_file
# import matplotlib.pyplot as plt
# %matplotlib inline
# import numpy as np
# c = SkyCoord(ra=wx*u.deg, dec=wy*u.deg, frame='icrs')
# c.galactic
import astropy.coordinates as coord
# import astropy.units as units
from astropy.coordinates import SkyCoord
from astropy.coordinates import Distance
from astropy.coordinates import Galactocentric
from astroquery.gaia import Gaia
Gaia.ROW_LIMIT = 10000  # Set the row limit for returned data

M45_center = SkyCoord.from_name('M45')
print(M45_center)

M45_distance = Distance.from_name('M45')
print(M45_distance)
print(M45_center.galactic)

print(M45_center.transform_to(Galactocentric))
gc1 = M45_center.transform_to(coord.Galactocentric)
