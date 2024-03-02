<template>
  <div class="file-upload">
    <div :style="{ display: 'flex', justifyContent: 'center' }">
      <a-card
        :style="{
          width: '360px',
          /* width: 100%; */
          border: ' 2px dashed #999',
          borderRadius: '20px',
          background: ' #efeccb',
          marginBottom: '15px',
        }"
        title="操作提示"
      >
        <template #extra>
          <a-link>More</a-link>
        </template>
        1.按顺序导入对应的文件<br />
        2.所有文件导入后，点击上传，铁塔文件数据量过大，请等待大概20分钟<br />
        3.导入的文件只需要保留需要获取数据的几列。<br />
        4.某些列数据如果为空，请统一改成0，不然后台数据无法处理
      </a-card>
    </div>
    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="总订单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success"> 总订单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="终止文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success"> 终止文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="详单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success"> 详单文件表变更</a-button>
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="铁塔账单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success">
            铁塔账单文件表变更</a-button
          >
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="铁塔订单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success">
            铁塔订单文件表变更</a-button
          >
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="室分账单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success">
            室分账单文件表变更</a-button
          >
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="微站账单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success">
            微站账单文件表变更</a-button
          >
        </a-space>
      </template>
    </a-upload>

    <a-upload
      :limit="1"
      :multiple="false"
      :auto-upload="false"
      @change="传输账单文件表变更;"
      :style="{
        marginBottom: '15px',
      }"
    >
      <template #upload-button>
        <a-space>
          <a-button type="dashed" status="success">
            传输账单文件表变更</a-button
          >
        </a-space>
      </template>
    </a-upload>
    <a-button
      :style="{
        borderRadius: '5px',
      }"
      size="large"
      @click="uploadFile"
      type="outline"
      status="success"
      >上传</a-button
    >
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
