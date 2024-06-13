export const domain = () => {
  return process.env.NODE_ENV ? `localhost:${process.env.PORT}` :'domain'
}