#!/usr/bin/env python3
import logging
import os
import mhi.pscad

# Log 'INFO' messages & above.  Include level & module name.
logging.basicConfig(level=logging.INFO,
                    format="%(levelname)-8s %(name)-26s %(message)s")

# Ignore INFO msgs from automation (eg, mhi.pscad, mhi.common, ...)
logging.getLogger('mhi').setLevel(logging.WARNING)

LOG = logging.getLogger('main')

# Launch PSCAD
pscad = mhi.pscad.launch(minimize=True)

if pscad:

    # Locate the tutorial directory
    tutorial_dir = os.path.join(pscad.examples_folder, "tutorial")
    LOG.info("Tutorial directory: %s", tutorial_dir)

    try:
        # Load the tutorial workspace
        pscad.load("Tutorial.pswx", folder=tutorial_dir)

        # Run all the simulation sets in the workspace
        pscad.run_all_simulation_sets()

    finally:
        # Exit PSCAD
        pscad.quit()

else:
    LOG.error("Failed to launch PSCAD")
