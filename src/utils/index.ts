export const numberToLetters = (num: number) => {
  num = num - 1
  let letters = ""
  while (num >= 0) {
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[num % 26] + letters
    num = Math.floor(num / 26) - 1
  }
  return letters
}

const lettersToNumber = (letters: string) => {
  return letters.split("").reduce((r, a) => r * 26 + parseInt(a, 36) - 9, 0)
}

export const addressToXY = (address: string) => {
  const regex = /([A-z]+)(\d+)/

  const match = address.match(regex)

  if (match === null) throw new Error()

  const col = lettersToNumber(match[1])
  const row = Number.parseInt(match[2])

  return [col, row] as const
}
