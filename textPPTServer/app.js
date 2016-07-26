var app = require('express')();
var server = require('http').Server(app);
var io = require('socket.io')(server);
var twilio = require('twilio');

const port = 3000

app.get('/', (request, response) => {  
  console.log("gotten")

  console.log(`server is listening on ${port}`)

  var resp = new twilio.TwimlResponse();
  resp.say('Testing Twilio and node.js');

   response.writeHead(200, {
        'Content-Type':'text/xml'
    });
    response.end(resp.toString());
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