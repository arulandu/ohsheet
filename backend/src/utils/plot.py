import numpy as np
from .conversion import address_to_coord
from api.models import SheetCacheData
import matplotlib.pyplot as plt

def plot_sheet_cache(cache: SheetCacheData, save_dir="./", save=False):
    scale = min(10 / cache.shape[0], 25 / (3*cache.shape[1]))
    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(3*cache.shape[1]*scale, cache.shape[0]*scale))
    shape = cache.shape

    plot_regions(cache.regions.format, shape, 'Format', ax1)
    plot_regions(cache.regions.formula, shape, 'Formula', ax2)
    plot_regions(cache.regions.color, shape, 'Color', ax3)

    if save:
        plt.savefig(f'{save_dir}/{cache.id}_regions.png')
    plt.show()

    if len(cache.info_ranges) > 0:
        fig, (ax1) = plt.subplots(1, 1, figsize=(4, 4))
        plot_ranges(cache.info_ranges, shape, 'Info Ranges', ax1)
        if save:
            plt.savefig(f'{save_dir}/{cache.id}_info_ranges.png')
        plt.show()

    if len(cache.tables) > 0:
        fig, (ax1) = plt.subplots(1, 1, figsize=(4, 4))
        plot_tables(cache.tables, shape, 'Tables', ax1)
        if save:
            plt.savefig(f'{save_dir}/{cache.id}_tables.png')
        plt.show()

def plot_tables(tables:list, shape, title, ax):
    colored_mask = np.zeros((*shape, 3))

    def color_range(rng:str, col):
        if ':' in rng:
            st, end = rng.split(':')
            (r1, c1), (r2, c2) = address_to_coord(st), address_to_coord(end)
            colored_mask[r1:r2+1, c1:c2+1, :] = col

    for i, table in enumerate(tables):
        col = np.random.random(3)*0.5+0.5
        color_range(table.data, col*0.5)
        color_range(table.row_hdr, col*0.9)
        color_range(table.col_hdr, col*1.1)
    
    colored_mask = np.clip(colored_mask, 0, 1)
    ax.imshow(colored_mask)
    ax.set_title(title)

def plot_ranges(rngs: dict, shape, title, ax):
    mask = np.zeros(shape)
    
    for i, rng in enumerate(rngs):
        st, end = rng.split(':')
        (r1, c1), (r2, c2) = address_to_coord(st), address_to_coord(end)
        mask[r1:r2+1, c1:c2+1] = i+1

    # Convert mask to colored image
    num_categories = len(rngs)
    colors = np.zeros((num_categories + 1, 3))
    colors[1:] = np.random.random((num_categories, 3))*0.5+0.5  # Random RGB values for categories 1+
    
    colored_mask = colors[mask.astype(int)]
    
    ax.imshow(colored_mask)
    ax.set_title(title)

def plot_regions(rngs: dict, shape, title, ax):
    mask = np.zeros(shape)
    
    # Assign random color to each range
    for i, region in enumerate(rngs):
        for rng in region['ranges']:
            st, end = rng.split(':')
            (r1, c1), (r2, c2) = address_to_coord(st), address_to_coord(end)
            mask[r1:r2+1, c1:c2+1] = i+1

    # Convert mask to colored image
    num_categories = int(mask.max())
    colors = np.zeros((num_categories + 1, 3))
    colors[1:] = np.random.random((num_categories, 3))*0.5+0.5  # Random RGB values for categories 1+
    
    colored_mask = colors[mask.astype(int)]
    
    ax.imshow(colored_mask)
    ax.set_title(title)