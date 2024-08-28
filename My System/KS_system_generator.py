"""From Starship Simulator"""
import random
import pandas as pd

def random_string(length=8):
    """Generate a random string of uppercase letters and digits."""
    return ''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=length))

def random_number(min_value, max_value):
    """Generate a random number within a specified range."""
    return random.randint(min_value, max_value)

def human_readable_star_name():
    """Generate a more interesting human-readable star name."""
    adjectives = ["Bright", "Distant", "Shining", "Glowing", "Mystic", "Ancient", "Celestial", "Radiant", "Silent", "Lonely"]
    nouns = ["Giant", "Dwarf", "Beacon", "Nova", "Star", "Light", "Phoenix", "Guardian", "Voyager", "Sentinel"]
    return f"{random.choice(adjectives)} {random.choice(nouns)}"

def random_sector():
    """Generate a random sector code."""
    return random_string(10)

# Function to generate the table
def generate_star_data_table(num_rows=100):
    rows = []
    unique_star_names = set()
    
    for i in range(num_rows):
        discord_name = random_string(6)
        while True:
            star_name = human_readable_star_name()
            if star_name not in unique_star_names:
                unique_star_names.add(star_name)
                break
        star_kelvin = random_number(3000, 30000)
        sector = random_sector()
        gal_x = random_number(-100000, 100000)
        gal_y = random_number(-100000, 100000)
        gal_z = random_number(-100000, 100000)
        desc = f"Star located in sector {sector}."
        
        rows.append([discord_name, star_name, star_kelvin, sector, gal_x, gal_y, gal_z, desc])

    columns = ["Discord Name", "Star Name", "Star Kelvin", "Sector", "Gal X", "Gal Y", "Gal Z", "Description"]
    df = pd.DataFrame(rows, columns=columns)
    
    return df

# Generate the table
df = generate_star_data_table(100)

# Display the table
print(df)
