// All the personal stuff lives here. Edit this file to change what the site says.

export const NAME = 'Kenu'
export const AGE = 23

export const surpriseMessage =
  'Happy birthday! (Your surprise message goes here.)'

// Photos: put image files in website/public/photos/ and use '/photos/your-file.jpg'.
export const memories: { date: string; title: string; text: string; photo?: string }[] = [
  { date: 'Month Year', title: 'The day we met', text: 'Write what happened here.' },
  { date: 'Month Year', title: 'That one trip', text: 'Another memory goes here.' },
  { date: 'Month Year', title: 'The time we laughed for an hour', text: 'And another one.' },
]

export const reasons: string[] = [
  'Reason #1 goes here',
  'An inside joke only you two get',
  'That nickname and where it came from',
  'Something only Kenu would do',
  'Another reason',
  'One more, because why not',
]

export const quiz: { question: string; options: string[]; answer: number }[] = [
  {
    question: "What is Kenu's favorite food?",
    options: ['Option A', 'Option B', 'Option C', 'Option D'],
    answer: 0,
  },
  {
    question: 'Where did we first meet?',
    options: ['Option A', 'Option B', 'Option C', 'Option D'],
    answer: 1,
  },
  {
    question: 'What does Kenu always say?',
    options: ['Option A', 'Option B', 'Option C', 'Option D'],
    answer: 2,
  },
]

export const favorites: { category: string; emoji: string; items: string[] }[] = [
  { category: 'Songs', emoji: '🎵', items: ['Song 1', 'Song 2', 'Song 3'] },
  { category: 'Food', emoji: '🍕', items: ['Food 1', 'Food 2'] },
  { category: 'Shows & Movies', emoji: '🎬', items: ['Show 1', 'Movie 1'] },
  { category: 'Random things', emoji: '✨', items: ['Thing 1', 'Thing 2'] },
]
