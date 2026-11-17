// A conservative candidate, never proof of association with a particular object.
export function isCaptionLabel(text){
  const value=text.trim()
  return value.length<=100 && /^(?:图|表)\s*\d+(?:[-－.]\d+)*(?:\s+|(?=[^\d.－-]))\S/u.test(value)
    && !/^(?:图|表)\s*\d+(?:[-－.]\d+)*\s*(?:为|是|显示|展示|所示)/u.test(value)
    && !/[。！？；]/u.test(value)
}
