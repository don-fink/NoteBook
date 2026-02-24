"""
Spell checking for QTextEdit with real-time underline highlighting and suggestions.
Uses pyenchant for dictionary lookups.
"""

import os
import re
from typing import List, Optional, Callable

from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QTextCursor, QTextCharFormat, QColor, QTextDocument

# Try to import enchant
try:
    import enchant
    from enchant.checker import SpellChecker
    ENCHANT_AVAILABLE = True
except ImportError:
    ENCHANT_AVAILABLE = False
    enchant = None
    SpellChecker = None


def _get_user_dictionary_path(language: str = "en_US") -> str:
    """Get path to user dictionary file in app data folder."""
    try:
        from settings_manager import get_app_data_dir
        data_dir = get_app_data_dir()
        os.makedirs(data_dir, exist_ok=True)
        return os.path.join(data_dir, f"user_dictionary_{language}.txt")
    except Exception:
        return os.path.join(os.path.expanduser("~"), f".notebook_dictionary_{language}.txt")


class SpellCheckHighlighter:
    """
    Manages spell checking for a QTextEdit widget.
    
    Uses QTextEdit's ExtraSelections to underline misspelled words with a red
    squiggly line. Provides right-click suggestions via context menu.
    """
    
    # Pattern to extract words (letters, apostrophes for contractions)
    WORD_PATTERN = re.compile(r"\b[A-Za-z']+\b")
    
    def __init__(self, text_edit: QtWidgets.QTextEdit, language: str = "en_US"):
        self._text_edit = text_edit
        self._language = language
        self._enabled = False
        self._dictionary = None  # enchant.DictWithPWL or None
        self._user_dict_path = _get_user_dictionary_path(language)
        self._check_timer = QTimer()
        self._check_timer.setSingleShot(True)
        self._check_timer.setInterval(300)  # Debounce spell check
        self._check_timer.timeout.connect(self._do_spell_check)
        self._original_context_menu_event = None
        self._misspelled_word_at_cursor = None  # Optional[str]
        self._misspelled_cursor_pos: int = 0
        
        # Underline format for misspelled words
        self._error_format = QTextCharFormat()
        self._error_format.setUnderlineStyle(QTextCharFormat.SpellCheckUnderline)
        self._error_format.setUnderlineColor(QColor(255, 0, 0))  # Red squiggly
        
        # Initialize dictionary with personal word list (PWL)
        if ENCHANT_AVAILABLE:
            try:
                # Create empty user dictionary file if it doesn't exist
                if not os.path.exists(self._user_dict_path):
                    os.makedirs(os.path.dirname(self._user_dict_path), exist_ok=True)
                    with open(self._user_dict_path, "w", encoding="utf-8") as f:
                        pass  # Create empty file
                # Use DictWithPWL for custom user dictionary
                self._dictionary = enchant.DictWithPWL(language, self._user_dict_path)
            except enchant.errors.DictNotFoundError:
                # Try fallback to en
                try:
                    self._user_dict_path = _get_user_dictionary_path("en")
                    if not os.path.exists(self._user_dict_path):
                        os.makedirs(os.path.dirname(self._user_dict_path), exist_ok=True)
                        with open(self._user_dict_path, "w", encoding="utf-8") as f:
                            pass
                    self._dictionary = enchant.DictWithPWL("en", self._user_dict_path)
                    self._language = "en"
                except Exception:
                    self._dictionary = None
            except Exception:
                self._dictionary = None
    
    @property
    def enabled(self) -> bool:
        return self._enabled
    
    @enabled.setter
    def enabled(self, value: bool):
        if value == self._enabled:
            return
        self._enabled = value
        if value:
            self._enable_spell_check()
        else:
            self._disable_spell_check()
    
    def _enable_spell_check(self):
        """Enable spell checking on the text edit."""
        if not ENCHANT_AVAILABLE or self._dictionary is None:
            return
        
        # Connect to text changes
        self._text_edit.textChanged.connect(self._on_text_changed)
        
        # Do initial spell check
        self._do_spell_check()
    
    def _disable_spell_check(self):
        """Disable spell checking and remove highlights."""
        try:
            self._text_edit.textChanged.disconnect(self._on_text_changed)
        except Exception:
            pass
        
        # Clear extra selections (only spell check ones)
        self._clear_spell_selections()
        
        # Stop pending checks
        self._check_timer.stop()
    
    def _on_text_changed(self):
        """Called when text changes - debounce spell check."""
        if self._enabled:
            self._check_timer.start()
    
    def _clear_spell_selections(self):
        """Clear spell check underlines without affecting other extra selections."""
        # Get current selections and filter out spell check ones
        # For now, we just clear all - in a more complex app we'd tag them
        try:
            self._text_edit.setExtraSelections([])
        except Exception:
            pass
    
    def _do_spell_check(self):
        """Perform spell check on the document."""
        if not self._enabled or self._dictionary is None:
            return
        
        try:
            doc = self._text_edit.document()
            text = doc.toPlainText()
            
            selections = []
            
            # Find all words and check spelling
            for match in self.WORD_PATTERN.finditer(text):
                word = match.group()
                
                # Skip very short words, numbers mixed with letters, etc.
                if len(word) < 2:
                    continue
                
                # Skip words that are all uppercase (acronyms)
                if word.isupper():
                    continue
                
                # Check spelling
                if not self._dictionary.check(word):
                    # Create selection for this word
                    cursor = self._text_edit.textCursor()
                    cursor.setPosition(match.start())
                    cursor.setPosition(match.end(), QTextCursor.KeepAnchor)
                    
                    selection = QtWidgets.QTextEdit.ExtraSelection()
                    selection.cursor = cursor
                    selection.format = self._error_format
                    selections.append(selection)
            
            self._text_edit.setExtraSelections(selections)
        except Exception:
            pass
    
    def _add_spell_suggestions_to_menu(self, menu, prepend=True, pos=None):
        """Add spell suggestions to an existing context menu.
        
        Args:
            menu: The context menu to add suggestions to
            prepend: If True (default), add suggestions at the top of the menu
            pos: Optional click position (QPoint) to determine which word was clicked
        """
        try:
            print(f"DEBUG _add_spell_suggestions: enabled={self._enabled}, dict={self._dictionary}")
            if not self._enabled or not self._dictionary:
                print("DEBUG: Returning early - not enabled or no dictionary")
                return
            
            # Get the word at the click position or current cursor position
            if pos is not None:
                # Get cursor at click position
                cursor = self._text_edit.cursorForPosition(pos)
                cursor.select(QTextCursor.WordUnderCursor)
            else:
                cursor = self._text_edit.textCursor()
                cursor.select(QTextCursor.WordUnderCursor)
            word = cursor.selectedText()
            print(f"DEBUG: word='{word}'")
            
            # Check if this word is misspelled
            if not word or len(word) < 2:
                print(f"DEBUG: Returning - word too short or empty")
                return
            is_correct = self._dictionary.check(word)
            print(f"DEBUG: is_correct={is_correct}")
            if is_correct:
                print("DEBUG: Word is spelled correctly, returning")
                return
            
            # Get suggestions
            suggestions = self._dictionary.suggest(word)[:5]  # Top 5 suggestions
            print(f"DEBUG: suggestions={suggestions}")
            
            # Insert suggestions at the top of the menu or append
            first_action = menu.actions()[0] if (prepend and menu.actions()) else None
            
            if suggestions:
                for suggestion in suggestions:
                    action = QtWidgets.QAction(suggestion, menu)
                    # Create a new cursor for each suggestion to avoid closure issues
                    word_start = cursor.selectionStart()
                    word_end = cursor.selectionEnd()
                    action.triggered.connect(
                        lambda checked, s=suggestion, ws=word_start, we=word_end: self._replace_word_at(ws, we, s)
                    )
                    action.setFont(QtGui.QFont(action.font().family(), -1, QtGui.QFont.Bold))
                    if prepend and first_action:
                        menu.insertAction(first_action, action)
                    else:
                        menu.addAction(action)
                
                # Add separator after suggestions
                if prepend and first_action:
                    menu.insertSeparator(first_action)
                else:
                    menu.addSeparator()
            
            # Add "Add to Dictionary" option
            add_action = QtWidgets.QAction(f'Add "{word}" to Dictionary', menu)
            add_action.triggered.connect(lambda checked, w=word: self._add_to_dictionary(w))
            if prepend and first_action:
                menu.insertAction(first_action, add_action)
            else:
                menu.addAction(add_action)
            
            # Add "Ignore" option
            ignore_action = QtWidgets.QAction(f'Ignore "{word}"', menu)
            ignore_action.triggered.connect(lambda checked, w=word: self._ignore_word(w))
            if prepend and first_action:
                menu.insertAction(first_action, ignore_action)
            else:
                menu.addAction(ignore_action)
            
            if prepend and first_action:
                menu.insertSeparator(first_action)
            else:
                menu.addSeparator()
        except Exception:
            pass
    
    def _replace_word_at(self, start: int, end: int, replacement: str):
        """Replace text between start and end with replacement."""
        try:
            cursor = self._text_edit.textCursor()
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.KeepAnchor)
            cursor.beginEditBlock()
            cursor.removeSelectedText()
            cursor.insertText(replacement)
            cursor.endEditBlock()
        except Exception:
            pass
    
    def _add_to_dictionary(self, word: str):
        """Add word to the personal dictionary (stored in app data folder)."""
        try:
            if self._dictionary:
                # add_to_pwl persists to the user dictionary file
                self._dictionary.add_to_pwl(word)
                # Re-check to remove underline
                self._do_spell_check()
        except Exception:
            pass
    
    def _ignore_word(self, word: str):
        """Ignore word for this session."""
        try:
            if self._dictionary:
                self._dictionary.add_to_session(word)
                # Re-check to remove underline
                self._do_spell_check()
        except Exception:
            pass
    
    def check_now(self):
        """Force an immediate spell check."""
        self._do_spell_check()
    
    def get_suggestions(self, word: str) -> List[str]:
        """Get spelling suggestions for a word."""
        if self._dictionary:
            return self._dictionary.suggest(word)
        return []
    
    def is_available(self) -> bool:
        """Check if spell checking is available."""
        return ENCHANT_AVAILABLE and self._dictionary is not None


def install_spell_check(text_edit: QtWidgets.QTextEdit, enabled: bool = True, language: str = "en_US") -> Optional[SpellCheckHighlighter]:
    """
    Install spell checking on a QTextEdit widget.
    
    Args:
        text_edit: The QTextEdit to add spell checking to
        enabled: Whether to enable immediately
        language: Language code (e.g., "en_US", "en_GB")
    
    Returns:
        SpellCheckHighlighter instance, or None if not available
    """
    if not ENCHANT_AVAILABLE:
        return None
    
    highlighter = SpellCheckHighlighter(text_edit, language)
    if not highlighter.is_available():
        return None
    
    highlighter.enabled = enabled
    
    # Store on widget for later access
    text_edit._spell_checker = highlighter
    
    return highlighter


def get_spell_checker(text_edit: QtWidgets.QTextEdit) -> Optional[SpellCheckHighlighter]:
    """Get the spell checker installed on a text edit, if any."""
    return getattr(text_edit, "_spell_checker", None)


def is_spell_check_available() -> bool:
    """Check if spell checking library is available."""
    return ENCHANT_AVAILABLE
