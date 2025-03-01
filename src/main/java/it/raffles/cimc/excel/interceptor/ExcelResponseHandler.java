package it.raffles.cimc.excel.interceptor;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.net.URLEncoder;
import java.util.List;

import org.springframework.core.MethodParameter;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.StringUtils;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;
import org.springframework.web.servlet.mvc.method.annotation.RequestResponseBodyMethodProcessor;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;

import it.raffles.cimc.excel.annotation.Excel;
import jakarta.servlet.http.HttpServletResponse;

public class ExcelResponseHandler implements HandlerMethodReturnValueHandler {

	private final RequestResponseBodyMethodProcessor target;

	public ExcelResponseHandler(RequestResponseBodyMethodProcessor target) {
		this.target = target;
	}

	@Override
	public boolean supportsReturnType(MethodParameter returnType) {
		return target.supportsReturnType(returnType);
	}

	@Override
	public void handleReturnValue(Object returnValue, MethodParameter returnType, ModelAndViewContainer mavContainer,
			NativeWebRequest webRequest) throws Exception {

		if (!returnType.hasMethodAnnotation(Excel.class)) {
			target.handleReturnValue(returnValue, returnType, mavContainer, webRequest);
			return;
		}

		if (!(returnValue instanceof List)) {
			target.handleReturnValue(returnValue, returnType, mavContainer, webRequest);
			return;
		}

		List<?> data = (List<?>) returnValue;
		Excel annotation = returnType.getMethodAnnotation(Excel.class);
		HttpServletResponse response = (HttpServletResponse) webRequest.getNativeResponse(HttpServletResponse.class);

		String charset = annotation.charset();
		String fileName = annotation.fileName();
		String contentType = annotation.contentType();
		String sheetName = annotation.sheetName();
		String password = annotation.password();

		ExcelTypeEnum suffix = annotation.suffix();
		String templateFileName = annotation.templateFileName();
		String templateFolderName = annotation.templateFolderName();

		boolean isLongestMatchColumnWidthStyleStrategy = annotation.isLongestMatchColumnWidthStyleStrategy();
		Class<?>[] entity = annotation.entity();
		Class<? extends WriteHandler>[] writeHandlers = annotation.writeHandlers();

		Class<?> clazz = 0 < entity.length ? entity[0]
				: (Class<?>) getSuperclassTypeParameter(returnType.getGenericParameterType());

		ExcelWriter writer = null;
		try {
			fileName = URLEncoder.encode(fileName, charset);
			response.setContentType(contentType);
			response.setCharacterEncoding(charset);
			response.setHeader("Content-disposition", String.format("attachment;filename=%s%s", fileName, suffix.getValue()));

			ExcelWriterBuilder builder = EasyExcel.write(response.getOutputStream(), clazz);

			if (StringUtils.hasText(password))
				builder = builder.password(password);

			if (isLongestMatchColumnWidthStyleStrategy)
				builder = builder.registerWriteHandler(new LongestMatchColumnWidthStyleStrategy());

			if (0 < writeHandlers.length)
				for (Class<? extends WriteHandler> handler : writeHandlers)
					try {
						builder = builder.registerWriteHandler(handler.newInstance());
					} catch (InstantiationException | IllegalAccessException e) {
						e.printStackTrace();
						throw new RuntimeException(e.getMessage());
					}

			if (StringUtils.hasText(templateFileName)) {
				writer = builder.withTemplate(getTemplate(templateFileName, templateFolderName)).autoCloseStream(false)
						.build();
				writer.fill(data, EasyExcel.writerSheet().build());
			} else {
				writer = builder.build();
				writer.write(data, EasyExcel.writerSheet(sheetName).build());
			}
		} catch (Exception exception) {
			exception.printStackTrace();
			response.reset();
			throw exception;

		} finally {
			if (null != writer)
				writer.finish();
		}
	}

	private InputStream getTemplate(String templateFileName, String templateFolderName) {

		String lowerCaseFileName = templateFileName.toLowerCase();

		if (!lowerCaseFileName.endsWith(ExcelTypeEnum.XLSX.getValue())
				&& !lowerCaseFileName.endsWith(ExcelTypeEnum.XLS.getValue()))
			templateFileName += ExcelTypeEnum.XLSX.getValue();

		ClassPathResource templateFile = new ClassPathResource(
				String.format("%s/%s", templateFolderName, templateFileName));
		try {
			return templateFile.getInputStream();
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("the excel template is not exsits");
		}
	}

	// TypeToken.getSuperclassTypeParameter()
	private static Type getSuperclassTypeParameter(Type superclass) {
		if (superclass instanceof Class) {
			throw new RuntimeException("Missing type parameter");
		}
		ParameterizedType parameterized = (ParameterizedType) superclass;
		return parameterized.getActualTypeArguments()[0];
	}
}
