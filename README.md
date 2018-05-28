# RipeConcepts: Production Associate - Foilbase Foil Image AI Saver
This script is used by [RipeConcepts](https://www.ripeconcepts.com) for saving AI files with different foil image for the foilbase standard of [Minted](https://www.minted.com).

## Usage
1. Open all colorways on AI.
1. Check the rules before doing the magick of the script:
    - Make sure all file names follow the standard format (e.g.: `MIN-123-IFS_A_FRT.ai`).
    - Make sure to tag all foil layers with `id:foil_artwork`.
    - Make sure that **"Paste Remembers Layers"** is checked under the layers' options.
1. Choose one colorway and add at least 3 foil images inside the clip group in the `id:foil_artwork` layer. Follow the order:
    - Gold
    - Rosegold
    - Silver
    - Gold Glitter
    - Silver Glitter
1. Browse and open the script under **File** > **Scripts** or just drag and drop the script directly to the AI.
1. Input the destination folder and wait until the script is done.
1. If errors occur, the script will abort its process and describe the error.
