class Sentence:
  """
  Utility parser to traverse and extract text fragments.
  """
  EOS = -1

  def __init__(self, value):
    """
    Creates a parser over a text.

    Args:
      value: Text to process.
    """
    self.value = value
    self.idx = 0

  def peek(self, value=0):
    """
    Returns the character at the current position + offset.

    Args:
      value: Relative offset from the current cursor.

    Returns:
      Character at the position or EOS if out of range.
    """
    if self.idx + value < len(self.value):
      return self.value[self.idx + value]
    return Sentence.EOS

  def next(self, value=1):
    """
    Advances the cursor N positions.

    Args:
      value: Number of positions to advance.
    """
    self.idx += value

  def peek_next(self):
    """
    Returns the current character and advances the cursor by 1.

    Returns:
      Current character or EOS if out of range.
    """
    ch = self.peek()
    self.next()
    return ch

  def is_eos(self):
    """
    Indicates whether the cursor is at the end of the text.

    Returns:
      True if no characters remain.
    """
    return self.idx >= len(self.value)


  def parse_until_word(self, word):
    """
    Extracts text until a word/token is found.

    Args:
      word: Expected ending token.

    Returns:
      Substring before the token or None if not found.
    """
    start = self.idx
    while not self.is_eos():
      if self.match(word):
        return self.value[start:self.idx - len(word)]
      self.next()
    return None


  def match(self, word):
    """
    Checks whether the text from the cursor matches a token.

    Args:
      word: Token to compare.

    Returns:
      True if it matches; advances the cursor in that case.
    """
    for f in range(len(word)):
      if self.peek(f) != word[f]:
        return False
    self.next(len(word))
    return True

  def substring(self, idx, idx2) -> str:
    """
    Returns a substring by absolute indices.

    Args:
      idx: Start index.
      idx2: End index (exclusive).

    Returns:
      Substring.
    """
    return self.value[idx:idx2]
