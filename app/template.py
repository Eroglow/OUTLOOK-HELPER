import re
from typing import Dict


PLACEHOLDER_PATTERN = re.compile(r"\{([^{}]+)\}")


def render_template(text: str, values: Dict[str, object], keep_unmatched: bool = True) -> str:
	"""
	Render a template by replacing {placeholders} with values from the mapping.

	- Double braces like {{ and }} are preserved by first escaping them.
	- If a key is missing:
		- keep_unmatched=True: leave it as {key}
		- keep_unmatched=False: replace with empty string
	"""
	if text is None:
		return ""

	# Protect escaped braces
	protected = text.replace("{{", "\u007B").replace("}}", "\u007D")

	def replace(match: re.Match) -> str:
		key = match.group(1).strip()
		if key in values and values[key] is not None:
			return str(values[key])
		return match.group(0) if keep_unmatched else ""

	result = PLACEHOLDER_PATTERN.sub(replace, protected)
	return result.replace("\u007B", "{").replace("\u007D", "}")


def render_subject_and_body(subject_template: str, body_template: str, values: Dict[str, object], html: bool) -> Dict[str, str]:
	subject = render_template(subject_template, values, keep_unmatched=True)
	body = render_template(body_template, values, keep_unmatched=True)
	# For HTML emails, leave body as-is; for plain text, ensure CRLFs are present
	return {"subject": subject, "body": body}
