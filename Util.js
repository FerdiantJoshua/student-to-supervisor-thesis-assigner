function getCol2Idx(idx2Col) {
    var col2Idx = {}
    for(let i = 0; i < idx2Col.length; i++)
      col2Idx[idx2Col[i]] = i
    return col2Idx;
  }
  