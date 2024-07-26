"""Convert ra / dec to Galactic"""
import astropy.units as u
import astropy.time
import astropy.coordinates

coord = astropy.coordinates.TETE(
    ra=25 * u.degree,
    dec=45 * u.degree,
    obstime=astropy.time.Time("2023-04-12")
)

coord = coord.transform_to(astropy.coordinates.Galactic())

print(coord.l.deg, coord.b.deg)
