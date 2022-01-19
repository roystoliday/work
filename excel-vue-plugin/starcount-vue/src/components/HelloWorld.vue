<template>
<div class="container">
  <h3>Hello World!</h3>
  <p>I'm from a component in Vue!</p>
  <p>But, still no cheddar.</p>
  <button @click="run()">Make cells yellow</button>
  <button @click="writeToLog()">Write to log</button>
</div>
</template>

<script>
export default {
  name: 'HelloWorld',
  methods: {
    async run() {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = "yellow";
        range.load("address");

        await context.sync();

        console.log(`The range address was "${range.address}".`);
      });
    },
    async tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
      }
    },
    writeToLog() {
      console.log('From HelloWorld, Hello.');
    }
  }
}

</script>

<style scoped>
.container {
  margin: 20px;
  padding: 5px;
  background-color: #ddd;
}
</style>