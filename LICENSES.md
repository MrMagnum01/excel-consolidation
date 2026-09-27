# Third-party licences

Every library this project installs, and the licence it ships under, as
read from each package's installed distribution metadata (`pip show`,
cross-checked against each package's own `License-Expression`/`Classifier`
metadata). All are open source; no closed-source or paid dependency is
used.

## Direct dependencies (pinned in `requirements.txt`)

Only these two are pinned. Everything else below is pulled in by them, and
the version shown is the one a clean install resolved, not a pin.

| Library | Pinned version | Licence | Used for |
|---|---|---|---|
| [openpyxl](https://openpyxl.readthedocs.io) | 3.1.5 | MIT | Reading and writing the `.xlsx` workbooks (generator and consolidator) |
| [pytest](https://docs.pytest.org/) | 9.1.1 | MIT | Test suite |

## Transitive dependencies (not pinned; versions from a clean install)

| Library | Resolved version | Licence | Pulled in by |
|---|---|---|---|
| [et_xmlfile](https://foss.heptapod.net/openpyxl/et_xmlfile) | 2.0.0 | MIT | openpyxl (its XML writer) |
| [packaging](https://pypi.org/project/packaging/) | 26.3 | Apache-2.0 OR BSD-2-Clause | pytest |
| [pluggy](https://pypi.org/project/pluggy/) | 1.6.0 | MIT | pytest |
| [iniconfig](https://pypi.org/project/iniconfig/) | 2.3.0 | MIT | pytest |
| [Pygments](https://pygments.org/) | 2.21.0 | BSD-2-Clause | pytest |

## Notices

These licences require their copyright and licence notices to be kept with
any copy or redistribution. This repository does not vendor or redistribute
any of these packages; they are installed from PyPI. Each installed package
carries its own licence and notice files, and those must be preserved in
any distribution that includes them.

No paid or closed-source service is used anywhere in this project.
