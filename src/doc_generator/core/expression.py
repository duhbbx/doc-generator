"""Expression parser and evaluator module."""

import re
from typing import Any

from simpleeval import simple_eval, EvalWithCompoundTypes


class ExpressionEvaluator:
    """Safe expression evaluator using simpleeval."""

    # Pattern to match placeholders like {{name}} or {{列名}}
    PLACEHOLDER_PATTERN = re.compile(r"\{\{([^}]+)\}\}")

    def __init__(self):
        """Initialize the expression evaluator."""
        self._functions = {
            # String functions
            "concat": self._concat,
            "upper": lambda s: str(s).upper(),
            "lower": lambda s: str(s).lower(),
            "strip": lambda s: str(s).strip(),
            "left": lambda s, n: str(s)[:int(n)],
            "right": lambda s, n: str(s)[-int(n):],
            "mid": lambda s, start, length: str(s)[int(start):int(start) + int(length)],
            "len": lambda s: len(str(s)),
            "replace": lambda s, old, new: str(s).replace(str(old), str(new)),

            # Math functions
            "sum": lambda *args: sum(self._to_number(a) for a in args),
            "avg": lambda *args: sum(self._to_number(a) for a in args) / len(args) if args else 0,
            "min": lambda *args: min(self._to_number(a) for a in args),
            "max": lambda *args: max(self._to_number(a) for a in args),
            "round": lambda x, n=0: round(self._to_number(x), int(n)),
            "abs": lambda x: abs(self._to_number(x)),
            "int": lambda x: int(self._to_number(x)),
            "float": lambda x: float(self._to_number(x)),

            # Conditional
            "if": lambda cond, true_val, false_val: true_val if cond else false_val,
            "ifempty": lambda val, default: default if val is None or str(val).strip() == "" else val,

            # Format functions
            "format": lambda fmt, *args: fmt.format(*args),
            "number_format": self._number_format,
        }

    @staticmethod
    def _concat(*args) -> str:
        """Concatenate multiple values into a string."""
        return "".join(str(a) if a is not None else "" for a in args)

    @staticmethod
    def _to_number(value: Any) -> float:
        """Convert value to number, handling None and strings."""
        if value is None:
            return 0
        if isinstance(value, (int, float)):
            return float(value)
        try:
            return float(str(value).replace(",", ""))
        except ValueError:
            return 0

    @staticmethod
    def _number_format(value: Any, decimal_places: int = 2, use_thousands_sep: bool = True) -> str:
        """Format a number with specified decimal places and thousands separator."""
        num = ExpressionEvaluator._to_number(value)
        if use_thousands_sep:
            return f"{num:,.{int(decimal_places)}f}"
        return f"{num:.{int(decimal_places)}f}"

    def extract_placeholders(self, expression: str) -> list[str]:
        """Extract all placeholder names from an expression.

        Args:
            expression: Expression string containing {{placeholders}}.

        Returns:
            List of placeholder names (without braces).
        """
        return self.PLACEHOLDER_PATTERN.findall(expression)

    def substitute_placeholders(self, expression: str, data: dict[str, Any]) -> str:
        """Replace placeholders with actual values for evaluation.

        Args:
            expression: Expression string with {{placeholders}}.
            data: Dictionary mapping placeholder names to values.

        Returns:
            Expression with placeholders replaced by values.
        """
        def replace(match):
            name = match.group(1)
            value = data.get(name)
            if value is None:
                return "None"
            if isinstance(value, str):
                # Escape quotes and wrap in quotes
                escaped = value.replace("\\", "\\\\").replace('"', '\\"')
                return f'"{escaped}"'
            return repr(value)

        return self.PLACEHOLDER_PATTERN.sub(replace, expression)

    def evaluate(self, expression: str, data: dict[str, Any]) -> Any:
        """Evaluate an expression with the given data.

        Args:
            expression: Expression string. Can be:
                - A simple placeholder: "{{name}}" -> returns value directly
                - An expression: "{{price}} * {{quantity}}" -> evaluates
                - A function call: "concat({{first}}, ' ', {{last}})"

        Returns:
            The evaluated result.
        """
        # Check if it's just a simple placeholder
        simple_match = re.fullmatch(r"\{\{([^}]+)\}\}", expression.strip())
        if simple_match:
            return data.get(simple_match.group(1))

        # Substitute placeholders and evaluate
        substituted = self.substitute_placeholders(expression, data)

        try:
            evaluator = EvalWithCompoundTypes(functions=self._functions)
            return evaluator.eval(substituted)
        except Exception as e:
            raise ValueError(f"Failed to evaluate expression '{expression}': {e}") from e

    def evaluate_safe(self, expression: str, data: dict[str, Any], default: Any = "") -> Any:
        """Evaluate an expression, returning default on error.

        Args:
            expression: Expression string.
            data: Dictionary mapping placeholder names to values.
            default: Value to return if evaluation fails.

        Returns:
            The evaluated result or default value.
        """
        try:
            result = self.evaluate(expression, data)
            return result if result is not None else default
        except Exception:
            return default


# Global evaluator instance
_evaluator = ExpressionEvaluator()


def evaluate_expression(expression: str, data: dict[str, Any]) -> Any:
    """Evaluate an expression with the given data.

    Convenience function using the global evaluator.
    """
    return _evaluator.evaluate(expression, data)


def extract_placeholders(expression: str) -> list[str]:
    """Extract placeholders from an expression.

    Convenience function using the global evaluator.
    """
    return _evaluator.extract_placeholders(expression)
