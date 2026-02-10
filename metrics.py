"""
Metrics and Validation Module

Provides functions for:
- Computing question usage statistics
- Calculating min/max/delta metrics (overall and per-difficulty)
- Validating quiz structure constraints
- Generating summary reports
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Set
from collections import defaultdict


def compute_usage_table(
    usage_counts: Dict[str, int],
    question_bank
) -> pd.DataFrame:
    """
    Create a DataFrame with question usage statistics.
    
    Args:
        usage_counts: Dict mapping question_id to total usage count
        question_bank: QuestionBank object containing question metadata
    
    Returns:
        DataFrame with columns: question_id, difficulty, total_usage_count
    """
    rows = []
    for q in question_bank.get_all():
        rows.append({
            'question_id': q.question_id,
            'difficulty': q.difficulty,
            'total_usage_count': usage_counts.get(q.question_id, 0)
        })
    
    df = pd.DataFrame(rows)
    # Sort by difficulty (hard, medium, easy) then by question_id
    difficulty_order = {'hard': 0, 'medium': 1, 'easy': 2}
    df['_sort'] = df['difficulty'].map(difficulty_order)
    df = df.sort_values(['_sort', 'question_id']).drop(columns=['_sort'])
    df = df.reset_index(drop=True)
    
    return df


def compute_min_max_delta(usage_counts: Dict[str, int]) -> Dict[str, int]:
    """
    Compute overall min, max, and delta of usage counts.
    
    Returns:
        Dict with keys: 'min_usage', 'max_usage', 'delta'
    """
    if not usage_counts:
        return {'min_usage': 0, 'max_usage': 0, 'delta': 0}
    
    counts = list(usage_counts.values())
    min_val = min(counts)
    max_val = max(counts)
    
    return {
        'min_usage': min_val,
        'max_usage': max_val,
        'delta': max_val - min_val
    }


def compute_difficulty_delta(
    usage_counts: Dict[str, int],
    question_bank
) -> pd.DataFrame:
    """
    Compute min, max, delta for each difficulty level.
    
    Returns:
        DataFrame with columns: difficulty, min, max, delta, variance
    """
    # Group usage counts by difficulty
    by_difficulty: Dict[str, List[int]] = defaultdict(list)
    for q in question_bank.get_all():
        count = usage_counts.get(q.question_id, 0)
        by_difficulty[q.difficulty].append(count)
    
    rows = []
    for difficulty in ['hard', 'medium', 'easy']:
        counts = by_difficulty.get(difficulty, [])
        if counts:
            min_val = min(counts)
            max_val = max(counts)
            variance = np.var(counts)
        else:
            min_val = max_val = 0
            variance = 0.0
        
        rows.append({
            'difficulty': difficulty,
            'min': min_val,
            'max': max_val,
            'delta': max_val - min_val,
            'variance': round(variance, 4)
        })
    
    return pd.DataFrame(rows)


def validate_quiz_structure(
    allocation_matrix: List[List[str]],
    question_bank
) -> Tuple[bool, List[str]]:
    """
    Validate that each student has valid questions (correct difficulty mapping).
    
    This is a dynamic validation - it checks that all questions exist and
    have valid difficulty values, but does not enforce specific counts.
    
    Args:
        allocation_matrix: 2D list of question assignments
        question_bank: QuestionBank for question metadata lookup
    
    Returns:
        Tuple of (is_valid, list of error messages)
    """
    errors = []
    
    # Build question_id -> difficulty mapping
    qid_to_diff = {q.question_id: q.difficulty for q in question_bank.get_all()}
    
    for student_idx, quiz in enumerate(allocation_matrix):
        student_label = f"S{student_idx + 1}"
        
        # Count difficulties in quiz
        diff_counts = {'hard': 0, 'medium': 0, 'easy': 0}
        for qid in quiz:
            diff = qid_to_diff.get(qid)
            if diff is None:
                errors.append(f"{student_label}: Unknown question ID: {qid}")
            elif diff in diff_counts:
                diff_counts[diff] += 1
            else:
                errors.append(f"{student_label}: Invalid difficulty '{diff}' for {qid}")
    
    return (len(errors) == 0, errors)


def validate_no_duplicates(
    allocation_matrix: List[List[str]]
) -> Tuple[bool, List[str]]:
    """
    Validate that no student has duplicate questions in their quiz.
    
    Returns:
        Tuple of (is_valid, list of error messages)
    """
    errors = []
    
    for student_idx, quiz in enumerate(allocation_matrix):
        seen: Set[str] = set()
        duplicates = []
        
        for qid in quiz:
            if qid in seen:
                duplicates.append(qid)
            seen.add(qid)
        
        if duplicates:
            student_label = f"S{student_idx + 1}"
            errors.append(
                f"{student_label}: Duplicate questions: {duplicates}"
            )
    
    return (len(errors) == 0, errors)


def validate_all_used(
    usage_counts: Dict[str, int],
    question_bank
) -> Tuple[bool, List[str]]:
    """
    Validate that all questions appear at least once.
    
    Returns:
        Tuple of (is_valid, list of unused questions)
    """
    unused = []
    
    for q in question_bank.get_all():
        if usage_counts.get(q.question_id, 0) == 0:
            unused.append(q.question_id)
    
    if unused:
        return (False, [f"Unused questions: {unused}"])
    
    return (True, [])


def run_all_validations(
    allocation_matrix: List[List[str]],
    usage_counts: Dict[str, int],
    question_bank
) -> Dict[str, Tuple[bool, List[str]]]:
    """
    Run all validation checks and return results.
    
    Returns:
        Dict mapping validation name to (passed, errors) tuple
    """
    return {
        'quiz_structure': validate_quiz_structure(allocation_matrix, question_bank),
        'no_duplicates': validate_no_duplicates(allocation_matrix),
        'all_questions_used': validate_all_used(usage_counts, question_bank),
    }


def print_validation_report(
    validations: Dict[str, Tuple[bool, List[str]]]
) -> bool:
    """
    Print a formatted validation report.
    
    Returns:
        True if all validations passed, False otherwise
    """
    print("\n" + "=" * 60)
    print("VALIDATION RESULTS")
    print("=" * 60)
    
    all_passed = True
    
    for name, (passed, errors) in validations.items():
        status = "✓ PASS" if passed else "✗ FAIL"
        print(f"\n{name}: {status}")
        
        if not passed:
            all_passed = False
            for error in errors[:5]:  # Limit to first 5 errors
                print(f"  - {error}")
            if len(errors) > 5:
                print(f"  ... and {len(errors) - 5} more errors")
    
    print("\n" + "-" * 60)
    overall = "ALL VALIDATIONS PASSED ✓" if all_passed else "SOME VALIDATIONS FAILED ✗"
    print(f"OVERALL: {overall}")
    print("=" * 60 + "\n")
    
    return all_passed


def generate_allocation_dataframe(
    allocation_matrix: List[List[str]]
) -> pd.DataFrame:
    """
    Convert allocation matrix to a DataFrame for display/export.
    
    Rows are quiz positions (Q1-Q15), columns are students (S1-S50).
    """
    num_students = len(allocation_matrix)
    num_positions = len(allocation_matrix[0]) if allocation_matrix else 0
    
    # Transpose: we want positions as rows, students as columns
    data = {}
    for student_idx in range(num_students):
        student_label = f"S{student_idx + 1}"
        data[student_label] = allocation_matrix[student_idx]
    
    df = pd.DataFrame(data)
    df.index = [f"Q{i + 1}" for i in range(num_positions)]
    
    return df


def compute_random_baseline_delta(
    question_bank,
    num_students: int = 50,
    num_trials: int = 100
) -> Dict[str, float]:
    """
    Compute expected delta from pure random allocation for comparison.
    
    Simulates random allocation multiple times and returns average metrics.
    """
    import random
    
    deltas = []
    
    for _ in range(num_trials):
        usage = defaultdict(int)
        
        for _ in range(num_students):
            assigned: Set[str] = set()
            
            # Hard questions (4)
            hard_qs = [q.question_id for q in question_bank.get_by_difficulty('hard')]
            for qid in random.sample(hard_qs, 4):
                usage[qid] += 1
                assigned.add(qid)
            
            # Medium questions (6)
            medium_qs = [q.question_id for q in question_bank.get_by_difficulty('medium')]
            for qid in random.sample(medium_qs, 6):
                usage[qid] += 1
            
            # Easy questions (5)
            easy_qs = [q.question_id for q in question_bank.get_by_difficulty('easy')]
            for qid in random.sample(easy_qs, 5):
                usage[qid] += 1
        
        counts = list(usage.values())
        if counts:
            deltas.append(max(counts) - min(counts))
    
    return {
        'avg_delta': round(np.mean(deltas), 2),
        'min_delta': min(deltas),
        'max_delta': max(deltas),
    }
