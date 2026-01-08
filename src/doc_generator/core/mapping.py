"""Mapping rules management module."""

import json
from dataclasses import dataclass, field, asdict
from enum import Enum
from pathlib import Path
from typing import Any


class MappingType(str, Enum):
    """Type of mapping rule."""
    DIRECT = "direct"  # Direct column to placeholder mapping
    EXPRESSION = "expression"  # Expression-based mapping


@dataclass
class MappingRule:
    """A single mapping rule."""
    placeholder: str  # Target placeholder name in Word template
    mapping_type: MappingType = MappingType.DIRECT
    source: str = ""  # Source column name (for direct mapping)
    expression: str = ""  # Expression string (for expression mapping)

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            "placeholder": self.placeholder,
            "type": self.mapping_type.value,
            "source": self.source,
            "expression": self.expression,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MappingRule":
        """Create from dictionary."""
        return cls(
            placeholder=data["placeholder"],
            mapping_type=MappingType(data.get("type", "direct")),
            source=data.get("source", ""),
            expression=data.get("expression", ""),
        )

    def get_expression(self) -> str:
        """Get the expression for evaluation.

        Returns:
            For direct mappings, returns "{{source}}".
            For expression mappings, returns the expression string.
        """
        if self.mapping_type == MappingType.DIRECT:
            return f"{{{{{self.source}}}}}" if self.source else ""
        return self.expression


@dataclass
class MappingConfig:
    """Complete mapping configuration."""
    rules: list[MappingRule] = field(default_factory=list)
    output_filename_pattern: str = "output_{{_index}}.docx"
    excel_file: str = ""
    template_file: str = ""
    output_directory: str = ""
    sheet_name: str = ""
    header_row: int = 1
    start_row: int = 2

    def to_dict(self) -> dict:
        """Convert to dictionary for JSON serialization."""
        return {
            "rules": [rule.to_dict() for rule in self.rules],
            "output_filename_pattern": self.output_filename_pattern,
            "excel_file": self.excel_file,
            "template_file": self.template_file,
            "output_directory": self.output_directory,
            "sheet_name": self.sheet_name,
            "header_row": self.header_row,
            "start_row": self.start_row,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MappingConfig":
        """Create from dictionary."""
        rules = [MappingRule.from_dict(r) for r in data.get("rules", [])]
        return cls(
            rules=rules,
            output_filename_pattern=data.get("output_filename_pattern", "output_{{_index}}.docx"),
            excel_file=data.get("excel_file", ""),
            template_file=data.get("template_file", ""),
            output_directory=data.get("output_directory", ""),
            sheet_name=data.get("sheet_name", ""),
            header_row=data.get("header_row", 1),
            start_row=data.get("start_row", 2),
        )

    def save(self, path: str | Path) -> None:
        """Save configuration to JSON file."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str | Path) -> "MappingConfig":
        """Load configuration from JSON file."""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return cls.from_dict(data)

    def get_mappings_dict(self) -> dict[str, str]:
        """Get mappings as a dictionary for word renderer.

        Returns:
            Dictionary mapping placeholder names to expressions.
        """
        return {rule.placeholder: rule.get_expression() for rule in self.rules if rule.placeholder}

    def add_rule(self, rule: MappingRule) -> None:
        """Add a mapping rule."""
        # Remove existing rule for same placeholder
        self.rules = [r for r in self.rules if r.placeholder != rule.placeholder]
        self.rules.append(rule)

    def remove_rule(self, placeholder: str) -> None:
        """Remove a mapping rule by placeholder name."""
        self.rules = [r for r in self.rules if r.placeholder != placeholder]

    def get_rule(self, placeholder: str) -> MappingRule | None:
        """Get a mapping rule by placeholder name."""
        for rule in self.rules:
            if rule.placeholder == placeholder:
                return rule
        return None

    def clear_rules(self) -> None:
        """Remove all mapping rules."""
        self.rules.clear()

    def auto_map(self, excel_columns: list[str], word_placeholders: list[str]) -> None:
        """Automatically create direct mappings for matching names.

        Args:
            excel_columns: List of column names from Excel.
            word_placeholders: List of placeholder names from Word template.
        """
        for placeholder in word_placeholders:
            if placeholder in excel_columns:
                self.add_rule(MappingRule(
                    placeholder=placeholder,
                    mapping_type=MappingType.DIRECT,
                    source=placeholder,
                ))
