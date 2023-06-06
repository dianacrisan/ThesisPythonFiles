#!/bin/bash
# Change directory
cd ~/Desktop/Dizertatie/CodeFiles
python3 automateDesignSmells.py
wait $!
python3 automateCommitComparisons.py
wait $!
python3 automateStatistics.py
wait $!
python3 automateEvaluation.py
wait $!

# Make the script executable using command 'chmod +x run_python_files.sh'
# Run the script by executing './run_python_files.sh'
