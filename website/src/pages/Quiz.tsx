import { useState } from 'react'
import { NAME, quiz } from '../content'
import './Quiz.css'

export default function Quiz() {
  const [step, setStep] = useState(0)
  const [score, setScore] = useState(0)
  const [picked, setPicked] = useState<number | null>(null)

  const done = step >= quiz.length
  const q = quiz[step]

  function pick(i: number) {
    if (picked !== null) return
    setPicked(i)
    if (i === q.answer) setScore((s) => s + 1)
    setTimeout(() => {
      setPicked(null)
      setStep((s) => s + 1)
    }, 900)
  }

  function restart() {
    setStep(0)
    setScore(0)
  }

  return (
    <main className="page quiz">
      <h1>How well do you know {NAME}?</h1>
      {done ? (
        <div className="screen">
          <p>GAME OVER</p>
          <p>SCORE: {score}/{quiz.length}</p>
          <button onClick={restart}>PLAY AGAIN</button>
        </div>
      ) : (
        <div className="screen">
          <p className="level">LEVEL {step + 1}/{quiz.length}</p>
          <p className="question">{q.question}</p>
          <div className="options">
            {q.options.map((o, i) => (
              <button
                key={i}
                onClick={() => pick(i)}
                className={picked === null ? '' : i === q.answer ? 'right' : i === picked ? 'wrong' : ''}
              >
                {o}
              </button>
            ))}
          </div>
        </div>
      )}
    </main>
  )
}
