import { useState } from 'react'
import { reasons } from '../content'
import './Reasons.css'

export default function Reasons() {
  const [flipped, setFlipped] = useState<Set<number>>(new Set())

  function flip(i: number) {
    setFlipped((prev) => new Set(prev).add(i))
  }

  return (
    <main className="page reasons">
      <h1>Reasons, jokes & things only you would do</h1>
      <div className="board">
        {reasons.map((r, i) => (
          <button key={i} className={`note${flipped.has(i) ? ' open' : ''}`} onClick={() => flip(i)}>
            {flipped.has(i) ? r : `#${i + 1}`}
          </button>
        ))}
      </div>
    </main>
  )
}
