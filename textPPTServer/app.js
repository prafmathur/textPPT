var app = require('express')();
var server = require('http').Server(app);
var io = require('socket.io')(server);
const port = 3000

app.get('/', (request, response) => {  
  console.log("gotten")
  response.send('Hello from Express!')
})

app.listen(port, (err) => {  
  if (err) {
    return console.log('something bad happened', err)
  }

  console.log(`server is listening on ${port}`)
})


io.on('connection', function (socket) {
  console.log("CONNECTED");
  socket.emit('news', { hello: 'world' });
  socket.on('my other event', function (data) {
    console.log(data);
  });
});