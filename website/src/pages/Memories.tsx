import { memories } from '../content'
import './Memories.css'

export default function Memories() {
  return (
    <main className="page memories">
      <h1>Our Memories</h1>
      <ol className="timeline">
        {memories.map((m, i) => (
          <li key={i} className="polaroid" style={{ rotate: `${i % 2 ? 2 : -2}deg` }}>
            {m.photo ? <img src={m.photo} alt={m.title} /> : <div className="photo-slot">photo</div>}
            <span className="date">{m.date}</span>
            <h2>{m.title}</h2>
            <p>{m.text}</p>
          </li>
        ))}
      </ol>
    </main>
  )
}
