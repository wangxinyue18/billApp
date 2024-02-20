<template>
  <div class="file-upload">
    <label>
      <input type="file" @change="handleFileChange" multiple />
      文件上传
    </label>

    <label>
      <input type="file" @change="handleFileChange" multiple />
      铁塔文件上传
    </label>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="总订单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 总订单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="终止文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 终止文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="详单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 详单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="铁塔账单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 铁塔账单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

     <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="铁塔订单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 铁塔订单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="室分账单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 室分账单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="微站账单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 微站账单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="传输账单文件表变更"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="primary"> 传输账单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>
    <button @click="uploadFile">上传</button>
  </div>
</template>

<script>
import { ipcRenderer } from "electron"; // 引入ipcRenderer模块
import { ref } from "vue";
export default {
  setup() {
    const 总订单文件表 = ref();
    const 终止文件表 = ref();
    const 铁塔账单文件表 = ref();
    const 铁塔订单文件表 = ref();
    const 室分账单文件表 = ref();
    const 微站账单文件表 = ref();
    const 传输账单文件表 = ref();
    const 详单文件表 = ref();
    const uploadFile = () => {
      const pathMap = {
        总订单文件表: 总订单文件表.value,
        铁塔账单文件表: 铁塔账单文件表.value,
        铁塔订单文件表: 铁塔订单文件表.value,
        终止文件表: 终止文件表.value,
        室分账单文件表: 室分账单文件表.value,
        微站账单文件表: 微站账单文件表.value,
        传输账单文件表: 传输账单文件表.value,
        详单文件表: 详单文件表.value,
      };

      // 使用IPC发送消息给主进程，让其处理文件上传
      console.log("🍎", pathMap);

      ipcRenderer.send("upload-file", JSON.stringify(pathMap));
    };
    const 总订单文件表变更 = (e) => {
      总订单文件表.value = e?.[0]?.file?.path;
      console.log(总订单文件表.value);
    };
    const 终止文件表变更 = (e) => (终止文件表.value = e?.[0]?.file?.path);
    const 铁塔账单文件表变更 = (e) =>
      (铁塔账单文件表.value = e?.[0]?.file?.path);
       const 铁塔订单文件表变更 = (e) =>
      (铁塔订单文件表.value = e?.[0]?.file?.path);
    const 室分账单文件表变更 = (e) =>
      (室分账单文件表.value = e?.[0]?.file?.path);
    const 微站账单文件表变更 = (e) =>
      (微站账单文件表.value = e?.[0]?.file?.path);
    const 传输账单文件表变更 = (e) =>
      (传输账单文件表.value = e?.[0]?.file?.path);
    const 详单文件表变更 = (e) => (详单文件表.value = e?.[0]?.file?.path);
    return {
      customRequest: console.log,
      总订单文件表变更,
      总订单文件表,
      终止文件表,
      终止文件表变更,
      铁塔账单文件表变更,
      铁塔账单文件表,
      铁塔订单文件表变更,
      铁塔订单文件表,
      室分账单文件表变更,
      室分账单文件表,
      微站账单文件表变更,
      微站账单文件表,
      传输账单文件表变更,
      传输账单文件表,
      详单文件表变更,
      详单文件表,
      uploadFile,
    };
  },
  methods: {
    handleFileChange(event) {
      const files = event.target.files;
      let arr = [];

      for (let i = 0; i < files.length; i++) {
        arr.push(files[i].path);
      }
      console.log("🦈", arr);

      // 使用IPC发送消息给主进程，让其处理文件上传

      ipcRenderer.send("upload-file", arr);
    },
  },
  data() {
    return {
      selectedFile: null,
    };
  },
};
</script>

<style scoped>
/* CSS样式 */
</style>
