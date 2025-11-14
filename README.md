#Market Fundamentals Processing Module

This repository contains a set of Python scripts used to process raw stock fundamentals and run straightforward financial calculations. Each script handles one part of the workflow, and main.py executes them in sequence.

get-mtrics.py extracts the numeric fields needed from the fundamental data and computes basic values like growth rates, margins, and simple ratios.
get-valuations.py applies standard valuation formulas directly on the cleaned numbers without any modeling or statistical layers.
cash-flow-earnings.py and Cash-flow-growth.py look at operating cash trends and multi-period cash flow changes using plain calculations with no abstraction.

All logic is written in regular Python with explicit formulas and visible steps. The goal is to take irregular financial statement data and convert it into consistent numeric outputs that can be reused in other analysis pipelines or screening workflows.
