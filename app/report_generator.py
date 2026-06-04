from pathlib import Path


class ReportGenerator:
    """Generate reports in PowerPoint and PDF formats."""

    def __init__(self, report_dir: str):
        self.report_dir = Path(report_dir)
        self.report_dir.mkdir(parents=True, exist_ok=True)

    def create_pptx(self, output_name: str) -> Path:
        """Create a PowerPoint report placeholder."""
        output_path = self.report_dir / output_name
        return output_path

    def create_pdf(self, output_name: str) -> Path:
        """Create a PDF report placeholder."""
        output_path = self.report_dir / output_name
        return output_path
