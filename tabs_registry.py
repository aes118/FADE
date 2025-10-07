from dataclasses import dataclass, field
from typing import Callable, Optional, List, Dict

@dataclass
class Tab:
    key: str
    label: str
    render: Callable[[], None]
    to_excel: Optional[Callable[[any], None]] = None
    from_excel: Optional[Callable[[any], None]] = None
    to_word: Optional[Callable[[any], None]] = None

@dataclass
class TabRegistry:
    ordered: List[str] = field(default_factory=list)
    by_key: Dict[str, Tab] = field(default_factory=dict)

    def register(self, tab: Tab, after: Optional[str] = None):
        """Add tab; if 'after' provided, insert after that key."""
        self.by_key[tab.key] = tab
        if after and after in self.ordered:
            self.ordered.insert(self.ordered.index(after) + 1, tab.key)
        else:
            self.ordered.append(tab.key)

    def labels(self) -> List[str]:
        return [self.by_key[k].label for k in self.ordered]

    def tabs(self) -> List[Tab]:
        return [self.by_key[k] for k in self.ordered]
