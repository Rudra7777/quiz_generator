"""
Answer Checking CLI — Part 2 of Quiz Generator

Two subcommands:
  generate  — Create a dummy student response sheet (simulated Google Form)
  check     — Validate responses and score students

Usage:
    python check_answers_main.py generate \
        --question-bank question_bank_72.xlsx \
        --question-papers output/question_papers.xlsx \
        --students 70 \
        --output output/student_responses.xlsx

    python check_answers_main.py check \
        --question-bank question_bank_72.xlsx \
        --question-papers output/question_papers.xlsx \
        --responses output/student_responses.xlsx \
        --output output/scoring_report.xlsx
"""

import argparse
from pathlib import Path

from excel_handler import load_question_bank
from response_generator import generate_responses, save_response_sheet
from answer_checker import (
    load_response_sheet,
    check_all_responses,
    generate_scoring_report,
)


def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Quiz Answer Checker — Part 2',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python check_answers_main.py generate --question-bank question_bank_72.xlsx --question-papers output/question_papers.xlsx --students 70
  python check_answers_main.py check --question-bank question_bank_72.xlsx --question-papers output/question_papers.xlsx --responses output/student_responses.xlsx
        '''
    )

    subparsers = parser.add_subparsers(dest='command', help='Command to run')

    # ── generate subcommand ──────────────────────────────────────────────
    gen_parser = subparsers.add_parser(
        'generate',
        help='Generate dummy student responses'
    )
    gen_parser.add_argument(
        '--question-bank', '-qb', required=True,
        help='Path to question bank Excel file'
    )
    gen_parser.add_argument(
        '--question-papers', '-qp', required=True,
        help='Path to generated question_papers.xlsx'
    )
    gen_parser.add_argument(
        '--students', '-s', type=int, required=True,
        help='Number of student responses to generate'
    )
    gen_parser.add_argument(
        '--output', '-o', default='output/student_responses.xlsx',
        help='Output path for response sheet (default: output/student_responses.xlsx)'
    )
    gen_parser.add_argument(
        '--correct-rate', type=float, default=0.70,
        help='Probability of correct answer (default: 0.70)'
    )
    gen_parser.add_argument(
        '--wrong-rate', type=float, default=0.20,
        help='Probability of wrong answer (default: 0.20)'
    )
    gen_parser.add_argument(
        '--seed', type=int, default=None,
        help='Random seed for reproducibility'
    )

    # ── check subcommand ─────────────────────────────────────────────────
    chk_parser = subparsers.add_parser(
        'check',
        help='Validate and score student responses'
    )
    chk_parser.add_argument(
        '--question-bank', '-qb', required=True,
        help='Path to question bank Excel file'
    )
    chk_parser.add_argument(
        '--question-papers', '-qp', required=True,
        help='Path to generated question_papers.xlsx'
    )
    chk_parser.add_argument(
        '--responses', '-r', required=True,
        help='Path to student responses Excel file'
    )
    chk_parser.add_argument(
        '--output', '-o', default='output/scoring_report.xlsx',
        help='Output path for scoring report (default: output/scoring_report.xlsx)'
    )

    return parser.parse_args()


def run_generate(args):
    """Generate dummy student responses."""
    print("=" * 70)
    print("ANSWER CHECKER — Generate Dummy Responses")
    print("=" * 70)

    # Validate inputs
    if not Path(args.question_bank).exists():
        print(f"\n❌ Error: Question bank not found: {args.question_bank}")
        return False
    if not Path(args.question_papers).exists():
        print(f"\n❌ Error: Question papers not found: {args.question_papers}")
        return False

    # Load question bank
    print(f"\n[1/3] Loading question bank: {args.question_bank}")
    question_bank = load_question_bank(args.question_bank)
    total = len(question_bank.get_all())
    print(f"  ✓ Loaded {total} questions")

    # Generate responses
    print(f"\n[2/3] Generating {args.students} student responses...")
    remaining_wrong = 1 - args.correct_rate - args.wrong_rate
    effective_wrong = args.wrong_rate + max(0.0, remaining_wrong)
    print(f"  Rates: {args.correct_rate:.0%} correct, {effective_wrong:.0%} wrong (compulsory answers)")

    response_df = generate_responses(
        question_papers_path=args.question_papers,
        question_bank=question_bank,
        num_students=args.students,
        correct_rate=args.correct_rate,
        wrong_rate=args.wrong_rate,
        blank_rate=0.0,
        seed=args.seed
    )

    # Save
    print(f"\n[3/3] Saving response sheet...")
    output_path = save_response_sheet(response_df, args.output)
    print(f"  ✓ Saved: {output_path}")

    # Summary
    num_cols = len(response_df.columns) - 1  # exclude Set_No
    print(f"\n{'=' * 70}")
    print(f"✅ Generated {len(response_df)} student responses")
    print(f"   Columns: Set_No + {num_cols} question columns = {num_cols + 1} total")
    print(f"   File: {output_path}")
    print(f"{'=' * 70}")

    return True


def run_check(args):
    """Validate and score student responses."""
    print("=" * 70)
    print("ANSWER CHECKER — Validate & Score")
    print("=" * 70)

    # Validate inputs
    for label, path in [
        ('Question bank', args.question_bank),
        ('Question papers', args.question_papers),
        ('Responses', args.responses),
    ]:
        if not Path(path).exists():
            print(f"\n❌ Error: {label} not found: {path}")
            return False

    # Load
    print(f"\n[1/4] Loading question bank: {args.question_bank}")
    question_bank = load_question_bank(args.question_bank)
    print(f"  ✓ {len(question_bank.get_all())} questions")

    print(f"\n[2/4] Loading student responses: {args.responses}")
    response_df = load_response_sheet(args.responses)
    print(f"  ✓ {len(response_df)} students, {len(response_df.columns)} columns")

    # Check
    print(f"\n[3/4] Validating and scoring...")
    report = check_all_responses(response_df, args.question_papers, question_bank)

    # Print summary
    print(f"\n  📊 Results:")
    print(f"     Students:       {len(report.student_reports)}")
    print(f"     Average Score:  {report.avg_score}%")
    print(f"     Median Score:   {report.median_score}%")
    print(f"     Pass Rate:      {report.pass_rate}% ({report.pass_count}/{len(report.student_reports)})")

    issues = report.validation_issues
    if issues:
        print(f"\n  ⚠️  Validation Issues: {len(issues)} students")
        for r in issues[:5]:
            print(f"     - Student {r.student_index + 1} ({r.validation.set_no}): "
                  f"{r.validation.extra_count} extra answers")
    else:
        print(f"\n  ✅ No validation issues — all students answered only their assigned questions")

    # Grade distribution
    grades = report.grade_distribution()
    print(f"\n  📈 Grade Distribution:")
    for grade, count in grades.items():
        bar = '█' * count
        print(f"     {grade:>10}: {count:3d}  {bar}")

    # Save report
    print(f"\n[4/4] Saving scoring report...")
    output_path = generate_scoring_report(report, args.output)
    print(f"  ✓ Saved: {output_path}")

    print(f"\n{'=' * 70}")
    print(f"✅ Scoring complete! Report: {output_path}")
    print(f"   Sheets: Scores, Summary, Validation")
    print(f"{'=' * 70}")

    return True


def main():
    """Main entry point."""
    args = parse_args()

    if args.command == 'generate':
        success = run_generate(args)
    elif args.command == 'check':
        success = run_check(args)
    else:
        print("Usage: python check_answers_main.py {generate|check} [options]")
        print("Run: python check_answers_main.py --help")
        return

    exit(0 if success else 1)


if __name__ == "__main__":
    main()
