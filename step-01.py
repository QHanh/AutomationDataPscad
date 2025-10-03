#!/usr/bin/env python3
import os
import mhi.pscad
import logging
logging.basicConfig(level=logging.INFO,
                    format="%(levelname)-8s %(name)-26s %(message)s")

# Launch PSCAD
pscad = mhi.pscad.launch(minimize=True)

# Locate the tutorial directory
tutorial_dir = os.path.join(pscad.examples_folder, "tutorial")

# Load the tutorial workspace
pscad.load("Tutorial.pswx", folder=tutorial_dir)

# Run all the simulation sets in the workspace
pscad.run_all_simulation_sets()

# Exit PSCAD
pscad.quit()
