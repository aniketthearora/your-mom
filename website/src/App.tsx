import { NavLink, Route, Routes } from 'react-router-dom'
import Home from './pages/Home'
import Memories from './pages/Memories'
import Reasons from './pages/Reasons'
import Quiz from './pages/Quiz'
import Favorites from './pages/Favorites'

const pages = [
  { path: '/', label: 'Home' },
  { path: '/memories', label: 'Memories' },
  { path: '/reasons', label: 'Reasons' },
  { path: '/quiz', label: 'Quiz' },
  { path: '/favorites', label: 'Favorites' },
]

export default function App() {
  return (
    <>
      <nav className="nav">
        {pages.map((p) => (
          <NavLink key={p.path} to={p.path} end>
            {p.label}
          </NavLink>
        ))}
      </nav>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/memories" element={<Memories />} />
        <Route path="/reasons" element={<Reasons />} />
        <Route path="/quiz" element={<Quiz />} />
        <Route path="/favorites" element={<Favorites />} />
        <Route path="*" element={<Home />} />
      </Routes>
    </>
  )
}
