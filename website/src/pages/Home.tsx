import { useState } from 'react'
import confetti from 'canvas-confetti'
import { AGE, NAME, surpriseMessage } from '../content'
import './Home.css'

export default function Home() {
  const [opened, setOpened] = useState(false)

  function open() {
    setOpened(true)
    confetti({ particleCount: 160, spread: 90, origin: { y: 0.6 } })
  }

  return (
    <main className="page home">
      <h1>Happy {AGE}rd, {NAME}!</h1>
      <p className="sub">This whole website is for you.</p>
      {opened ? (
        <p className="surprise">{surpriseMessage}</p>
      ) : (
        <button className="gift" onClick={open} aria-label="Open your gift">
          🎁
          <span>tap to open</span>
        </button>
      )}
    </main>
  )
}
