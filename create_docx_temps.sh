#!/bin/bash

# Set the base directory
BASE_DIR="$HOME/Desktop/docx"

# Create the base docx folder if it doesn't exist
mkdir -p "$BASE_DIR"

# Array of different Lorem Ipsum texts (10 unique snippets)
LOREM[1]="Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
LOREM[2]="Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
LOREM[3]="Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
LOREM[4]="Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
LOREM[5]="Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium."
LOREM[6]="Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit."
LOREM[7]="Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit."
LOREM[8]="Ut aut reiciendis voluptatibus maiores alias consequatur aut perferendis doloribus asperiores repellat."
LOREM[9]="Temporibus autem quibusdam et aut officiis debitis aut rerum necessitatibus saepe eveniet."
LOREM[10]="At vero eos et accusamus et iusto odio dignissimos ducimus qui blanditiis praesentium."

# Loop to create 10 subfolders and .docx files
for i in {1..10}
do
    # Create subfolder
    SUBFOLDER="$BASE_DIR/subfolder$i"
    mkdir -p "$SUBFOLDER"
    
    # Create a temporary markdown file
    echo "# Document $i" > temp.md
    echo "" >> temp.md
    echo "## Summary" >> temp.md
    echo "${LOREM[$i]}" >> temp.md
    
    # Convert markdown to docx using pandoc
    pandoc temp.md -o "$SUBFOLDER/document$i.docx"
    
    echo "Created document$i.docx in $SUBFOLDER"
done

# Clean up temporary markdown file
rm temp.md
