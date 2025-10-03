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

versions = mhi.pscad.versions()
LOG.info("PSCAD Versions: %s", versions)

# Skip any 'Beta' versions
versions = [(ver, x64) for ver, x64 in versions if "Beta" not in ver]

# Skip versions prior to '5.0'
versions = [(ver, x64) for ver, x64 in versions if not ver.startswith("4.")]

# Skip any 32-bit versions, if other choices exist
vers = [(ver, x64) for ver, x64 in versions if x64]
if len(vers) > 0:
    versions = vers

LOG.info("   After filtering: %s", versions)

# Of any remaining versions, choose the "lexically largest" one.
version, x64 = sorted(versions)[-1]
LOG.info("   Selected PSCAD version: %s %d-bit", version, 64 if x64 else 32)

# Launch PSCAD
LOG.info("Launching: %s", version)
pscad = mhi.pscad.launch(minimize=True, version=version, x64=x64)

if pscad:

    # Create a dictionary of new settings to apply
    new_settings = {}

    # Get Fortran compiler versions
    fortrans = pscad.setting_range('fortran_version')
    LOG.info("Fortran versions: %s", fortrans)

    # Skip 'GFortran' compilers, if other choices exist
    vers = [ver for ver in fortrans if 'GFortran' not in ver]
    if len(vers) > 0:
        fortrans = vers

    LOG.info("   After filtering: %s", fortrans)

    # Order the remaining compilers, choose the last one (highest revision)
    fortran = sorted(fortrans)[-1]
    new_settings['fortran_version'] = fortran
    LOG.info("   Selected Fortran version: %r", fortran)

    if pscad.version_number >= (5, 1) and 'Intel' in fortran:
        # In PSCAD 5.1+ w/ Intel Fortran compiler, a Visual Studios linker
        # selection required.
        linkers = pscad.setting_range('c_version')
        LOG.info("Linker versions: %s", linkers)
        linkers = {ver for ver in linkers if ver.startswith('VS')}
        LOG.info("   After filtering: %s", linkers)
        linker = sorted(linkers)[0]
        new_settings['c_version'] = linker
        LOG.info("   Selected Linker version: %r", linker)

    pscad.settings(**new_settings)

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
