#!/usr/bin/env python3
"""Test script for the schedule solver without GUI."""

import sys
import os
sys.path.append(os.path.dirname(__file__))

# Import solver
import importlib.util
solver_path = os.path.join(os.path.dirname(__file__), "solver.py.py")
spec = importlib.util.spec_from_file_location("solver", solver_path)
solver_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(solver_module)
ScheduleOptimizer = solver_module.ScheduleOptimizer

def test_solver():
    """Test the solver with preferences."""
    print("Testing ScheduleOptimizer with preferences...")

    # Test preferences
    preferences = {0: [0], 1: [1], 2: [0, 1]}  # employee 0 prefers shift 0, etc.

    # Create optimizer
    optimizer = ScheduleOptimizer(num_people=14, num_days=5, shifts_per_day=2, preferences=preferences)

    # Solve
    result = optimizer.solve()

    print(f"Status: {result['status']}")
    print(f"Schedule shape: {len(result['schedule'])}x{len(result['schedule'][0])}")

    # Count assignments per employee
    assignments = {}
    for p in range(len(result['schedule'])):
        count = 0
        for d in range(len(result['schedule'][p])):
            if result['schedule'][p][d]:
                count += 1
        assignments[p] = count

    print(f"Assignments per employee: {assignments}")

    # Check preferences
    preference_count = 0
    for p, prefs in preferences.items():
        if p < len(result['schedule']):
            for d in range(len(result['schedule'][p])):
                cell = result['schedule'][p][d]
                if cell and isinstance(cell, dict):
                    shift_name = cell.get('value', '')
                    if shift_name == 'Poranna' and 0 in prefs:
                        preference_count += 1
                    elif shift_name == 'Popołudniowa' and 1 in prefs:
                        preference_count += 1

    print(f"Preference assignments: {preference_count}")
    print("Test completed successfully!")

if __name__ == "__main__":
    test_solver()