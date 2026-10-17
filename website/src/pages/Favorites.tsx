import { favorites, NAME } from '../content'
import './Favorites.css'

export default function Favorites() {
  return (
    <main className="page favorites">
      <h1>{NAME}'s Favorite Things</h1>
      <div className="grid">
        {favorites.map((f) => (
          <section key={f.category} className="card">
            <h2>
              <span>{f.emoji}</span> {f.category}
            </h2>
            <ul>
              {f.items.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
          </section>
        ))}
      </div>
    </main>
  )
}
