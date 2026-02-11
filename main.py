"""
Quiz Question Allocation Engine - Main Entry Point

Dynamic quiz generation system that:
- Loads question bank from Excel file
- Allocates questions using greedy load-balancing
- Generates individual question papers for each student

Usage:
    python main.py --input question_bank.xlsx --students 50 --hard 4 --medium 6 --easy 5

Or with defaults:
    python main.py --input question_bank.xlsx --students 50
"""

import os
import argparse
import pandas as pd
from pathlib import Path

from allocator import (
    QuizStructure,
    allocate_quizzes,
    shuffle_all_quizzes,
)
from excel_handler import (
    load_question_bank,
    generate_question_papers,
    create_sample_question_bank_excel,
    FullQuestionBank,
)
from metrics import (
    compute_min_max_delta,
    compute_difficulty_delta,
    run_all_validations,
    print_validation_report,
    generate_allocation_dataframe,
)


# Default configuration
DEFAULT_HARD = 4
DEFAULT_MEDIUM = 6
DEFAULT_EASY = 5
DEFAULT_STUDENTS = 50


def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Quiz Question Allocation Engine',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python main.py --input question_bank.xlsx --students 50
  python main.py --input questions.xlsx --students 30 --hard 5 --medium 8 --easy 7
  python main.py --create-sample  # Creates a sample question bank
        '''
    )
    
    parser.add_argument(
        '--input', '-i',
        type=str,
        help='Path to input Excel file with question bank'
    )
    
    parser.add_argument(
        '--students', '-s',
        type=int,
        default=DEFAULT_STUDENTS,
        help=f'Number of students (default: {DEFAULT_STUDENTS})'
    )
    
    parser.add_argument(
        '--hard', '-H',
        type=int,
        default=DEFAULT_HARD,
        help=f'Number of hard questions per quiz (default: {DEFAULT_HARD})'
    )
    
    parser.add_argument(
        '--medium', '-M',
        type=int,
        default=DEFAULT_MEDIUM,
        help=f'Number of medium questions per quiz (default: {DEFAULT_MEDIUM})'
    )
    
    parser.add_argument(
        '--easy', '-E',
        type=int,
        default=DEFAULT_EASY,
        help=f'Number of easy questions per quiz (default: {DEFAULT_EASY})'
    )
    
    parser.add_argument(
        '--output', '-o',
        type=str,
        default='output',
        help='Output directory (default: output)'
    )
    
    parser.add_argument(
        '--create-sample',
        action='store_true',
        help='Create a sample question bank Excel file'
    )
    
    parser.add_argument(
        '--seed',
        type=int,
        default=None,
        help='Optional random seed for reproducible output (default: random each run)'
    )
    
    return parser.parse_args()


def create_simple_question_bank(qb: FullQuestionBank):
    """
    Create a simple QuestionBank-like interface for metrics module.
    Returns an object that works with the existing metrics functions.
    """
    class SimpleQB:
        def __init__(self, full_qb: FullQuestionBank):
            self._qb = full_qb
        
        def get_all(self):
            class SimpleQ:
                def __init__(self, q):
                    self.question_id = q.question_id
                    self.difficulty = q.difficulty
            return [SimpleQ(q) for q in self._qb.get_all()]
        
        def get_by_difficulty(self, difficulty):
            return self._qb.get_by_difficulty(difficulty)
        
        def count_by_difficulty(self):
            return self._qb.count_by_difficulty()
    
    return SimpleQB(qb)


def main():
    """Main entry point."""
    args = parse_args()
    
    # Ensure output directory exists
    os.makedirs(args.output, exist_ok=True)
    
    print("=" * 70)
    print("QUIZ QUESTION ALLOCATION ENGINE")
    print("=" * 70)
    
    # ========================================================================
    # Handle --create-sample flag
    # ========================================================================
    if args.create_sample:
        sample_path = "sample_question_bank.xlsx"
        print(f"\n📝 Creating sample question bank: {sample_path}")
        create_sample_question_bank_excel(
            sample_path,
            hard_count=10,
            medium_count=25,
            easy_count=15
        )
        print(f"✓ Sample file created with 50 questions (10H, 25M, 15E)")
        print(f"\nRun: python main.py --input {sample_path} --students 50")
        return True
    
    # ========================================================================
    # Validate input file
    # ========================================================================
    if not args.input:
        print("\n❌ Error: --input is required (or use --create-sample)")
        print("Run: python main.py --help")
        return False
    
    if not Path(args.input).exists():
        print(f"\n❌ Error: Input file not found: {args.input}")
        return False
    
    # ========================================================================
    # Step 1: Load Question Bank
    # ========================================================================
    print(f"\n[1/5] Loading question bank from: {args.input}")
    
    question_bank = load_question_bank(args.input)
    counts = question_bank.count_by_difficulty()
    
    print(f"  Question Bank Loaded:")
    print(f"    - Hard:   {counts.get('hard', 0)} questions")
    print(f"    - Medium: {counts.get('medium', 0)} questions")
    print(f"    - Easy:   {counts.get('easy', 0)} questions")
    print(f"    - Total:  {len(question_bank.get_all())} questions")
    
    # ========================================================================
    # Step 2: Configure Quiz Structure
    # ========================================================================
    quiz_structure = QuizStructure(
        hard_count=args.hard,
        medium_count=args.medium,
        easy_count=args.easy
    )
    
    print(f"\n[2/5] Quiz Configuration:")
    print(f"    - Students:         {args.students}")
    print(f"    - Questions/quiz:   {quiz_structure.total_questions()}")
    print(f"    - Structure:        {args.hard}H + {args.medium}M + {args.easy}E")
    
    # Validate configuration
    valid, errors = quiz_structure.validate(counts)
    if not valid:
        print(f"\n❌ Configuration Error:")
        for e in errors:
            print(f"    - {e}")
        return False
    
    # ========================================================================
    # Step 3: Run Allocation
    # ========================================================================
    print(f"\n[3/5] Allocating questions to {args.students} students...")
    print("  Using: Greedy Load-Balancing with Min-Heap")
    if args.seed is None:
        print("  Randomization: fresh each run (no fixed seed)")
    else:
        print(f"  Randomization: fixed seed = {args.seed}")
    
    # Get question IDs by difficulty
    q_ids_by_diff = {
        'hard': question_bank.get_question_ids_by_difficulty('hard'),
        'medium': question_bank.get_question_ids_by_difficulty('medium'),
        'easy': question_bank.get_question_ids_by_difficulty('easy'),
    }
    
    allocation_matrix, usage_counts = allocate_quizzes(
        q_ids_by_diff,
        num_students=args.students,
        quiz_structure=quiz_structure,
        seed=args.seed
    )
    
    print(f"  ✓ Generated {len(allocation_matrix)} quizzes")
    
    # Shuffle questions within each quiz
    print("  ✓ Shuffling question order...")
    shuffled_matrix = shuffle_all_quizzes(allocation_matrix, base_seed=args.seed)
    
    # ========================================================================
    # Step 4: Generate Output Files
    # ========================================================================
    print(f"\n[4/5] Generating output files...")
    
    # 4a. Question Papers (multi-sheet Excel)
    papers_path = os.path.join(args.output, "question_papers.xlsx")
    generate_question_papers(shuffled_matrix, question_bank, papers_path)
    print(f"  ✓ Saved: {papers_path}")
    
    # 4b. Allocation Table (original order)
    allocation_df = generate_allocation_dataframe(allocation_matrix)
    allocation_df.to_csv(os.path.join(args.output, "allocation_table.csv"))
    print(f"  ✓ Saved: {args.output}/allocation_table.csv")
    
    # 4c. Shuffled Table
    shuffled_df = generate_allocation_dataframe(shuffled_matrix)
    shuffled_df.to_csv(os.path.join(args.output, "shuffled_table.csv"))
    print(f"  ✓ Saved: {args.output}/shuffled_table.csv")
    
    # 4d. Evaluation Summary
    simple_qb = create_simple_question_bank(question_bank)
    overall_stats = compute_min_max_delta(usage_counts)
    difficulty_stats = compute_difficulty_delta(usage_counts, simple_qb)
    
    # Combined evaluation
    overall_row = pd.DataFrame([{
        'difficulty': 'OVERALL',
        'min': overall_stats['min_usage'],
        'max': overall_stats['max_usage'],
        'delta': overall_stats['delta'],
        'variance': '-'
    }])
    combined_eval = pd.concat([difficulty_stats, overall_row], ignore_index=True)
    combined_eval.to_csv(os.path.join(args.output, "evaluation.csv"), index=False)
    print(f"  ✓ Saved: {args.output}/evaluation.csv")
    
    # ========================================================================
    # Step 5: Display Results
    # ========================================================================
    print(f"\n[5/5] Results Summary")
    
    print("\n" + "-" * 70)
    print("EVALUATION METRICS")
    print("-" * 70)
    print(combined_eval.to_string(index=False))
    
    # Run validations
    validations = run_all_validations(allocation_matrix, usage_counts, simple_qb)
    all_passed = print_validation_report(validations)
    
    # ========================================================================
    # Summary
    # ========================================================================
    print("\n" + "=" * 70)
    if all_passed:
        print("✅ SUCCESS: All constraints satisfied!")
    else:
        print("⚠️  WARNING: Some validations failed")
    print("=" * 70)
    
    print("\n📁 OUTPUT FILES:")
    print(f"  - {args.output}/question_papers.xlsx  (Question papers for all students)")
    print(f"  - {args.output}/allocation_table.csv  (Original allocation)")
    print(f"  - {args.output}/shuffled_table.csv    (Shuffled allocation)")
    print(f"  - {args.output}/evaluation.csv        (Metrics summary)")
    
    print(f"\n📊 The question_papers.xlsx contains {args.students} sheets (Set_1 to Set_{args.students})")
    print("   Each sheet has questions WITHOUT answers for students.")
    print("   The 'Answer_Key' sheet contains all answers for teachers.")
    
    return all_passed


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
